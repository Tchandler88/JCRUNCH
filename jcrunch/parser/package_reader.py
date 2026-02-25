import zipfile
import os
import tempfile

from parser.xml_parser import parse_content_xml


def walk_package(zip_path: str) -> dict:
    """
    Unzip AEM package, walk jcr_root/ recursively.
    Parse every .content.xml and collect into in-memory harvest dict.
    No database. No file writes beyond a single reused temp file.
    Returns harvest dict only.

    harvest = {
        'nodes':          {},   # keyed by jcr_path
        'properties':     {},   # keyed by (jcr_path, full_name)
        'tags':           {},   # keyed by tag_id
        'tag_assignments': [],  # list of {jcr_path, tag_path} dicts
        'namespaces':     {},   # keyed by namespace URI
        'folders':        {},   # keyed by folder_path
    }

    Windows note: zipfile.extractall() fails on paths containing colons
    (e.g. cq:tags). We read each entry's bytes directly from the zip with
    zf.read() and write to a single safe temp file before parsing.
    The JCR path is derived from the zip entry name string, not the
    filesystem path — so colons in AEM paths are never a problem.
    """
    harvest = {
        'nodes':           {},
        'properties':      {},
        'tags':            {},
        'tag_assignments': [],
        'namespaces':      {},
        'folders':         {},
    }

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_xml = os.path.join(tmpdir, '_content.xml')

        with zipfile.ZipFile(zip_path, 'r') as zf:
            all_entries = zf.namelist()

            # Locate jcr_root prefix inside the zip entry names
            jcr_prefix = None
            for entry in all_entries:
                normalized = entry.replace('\\', '/')
                parts = normalized.split('/')
                if 'jcr_root' in parts:
                    idx = parts.index('jcr_root')
                    jcr_prefix = '/'.join(parts[:idx + 1]) + '/'
                    break

            if not jcr_prefix:
                raise ValueError(
                    "No jcr_root/ found — "
                    "is this a valid AEM Package Manager export?"
                )

            for zip_entry in sorted(all_entries):
                normalized = zip_entry.replace('\\', '/')

                if not normalized.startswith(jcr_prefix):
                    continue

                parts = normalized.split('/')
                if parts[-1] != '.content.xml':
                    continue

                # Build JCR path from zip entry name (no filesystem colon issue)
                rel = normalized[len(jcr_prefix):]
                if '/' in rel:
                    dir_part = rel.rsplit('/', 1)[0]
                    jcr_path = '/' + dir_part
                else:
                    jcr_path = '/'

                # AEM folder notation: _jcr_content → jcr:content
                jcr_path = jcr_path.replace('/_jcr_content', '/jcr:content')

                # Write to safe temp file — avoids Windows colon restriction
                with open(tmp_xml, 'wb') as f:
                    f.write(zf.read(zip_entry))

                try:
                    result = parse_content_xml(tmp_xml, jcr_path)
                    if not result:
                        continue

                    # Store node — dict deduplicates by path
                    # last write wins on re-run (idempotent)
                    harvest['nodes'][jcr_path] = {
                        'path':             jcr_path,
                        'node_type':        result.get('node_type'),
                        'resource_type':    result.get('resource_type'),
                        'template':         result.get('template'),
                        'last_modified':    result.get('last_modified'),
                        'last_modified_by': result.get('last_modified_by'),
                    }

                    # Store properties — keyed by (path, full_name)
                    for prop in result.get('properties', []):
                        key = (jcr_path, prop['full_name'])
                        harvest['properties'][key] = {
                            'jcr_path':  jcr_path,
                            'namespace': prop.get('namespace', ''),
                            'name':      prop.get('name'),
                            'full_name': prop.get('full_name'),
                            'value':     prop.get('value'),
                            'is_multi':  prop.get('is_multi', False),
                        }

                    # Store tag assignments as list
                    for tag_path in result.get('tags', []):
                        harvest['tag_assignments'].append({
                            'jcr_path': jcr_path,
                            'tag_path': tag_path,
                        })

                    # Store namespaces — keyed by URI
                    for prefix, uri in result.get('namespaces', {}).items():
                        if uri not in harvest['namespaces']:
                            harvest['namespaces'][uri] = {
                                'uri':    uri,
                                'prefix': prefix,
                            }

                    # Store folder — keyed by path
                    folder_path = _extract_folder_path(jcr_path)
                    if folder_path and folder_path not in harvest['folders']:
                        harvest['folders'][folder_path] = {
                            'folder_path':   folder_path,
                            'folder_name':   folder_path.rsplit('/', 1)[-1],
                            'depth_level':   folder_path.count('/'),
                            'parent_folder': (
                                folder_path.rsplit('/', 1)[0]
                                if '/' in folder_path.lstrip('/')
                                else ''
                            ),
                        }

                    # If this is a tag definition node, store it
                    if '/content/cq:tags/' in jcr_path:
                        tag_id = jcr_path.replace(
                            '/content/cq:tags/', ''
                        ).strip('/')
                        if tag_id:
                            title_key = (jcr_path, 'jcr:title')
                            desc_key  = (jcr_path, 'jcr:description')
                            title = harvest['properties'].get(
                                title_key, {}
                            ).get('value', '')
                            desc  = harvest['properties'].get(
                                desc_key, {}
                            ).get('value', '')
                            harvest['tags'][tag_id] = {
                                'tag_id':      tag_id,
                                'tag_title':   title,
                                'description': desc,
                                'asset_count': 0,
                            }

                except Exception as e:
                    print(f"   WARNING Skipping {jcr_path}: {e}")
                    continue

    # Count tag usage from tag_assignments
    for assignment in harvest['tag_assignments']:
        raw    = assignment['tag_path']
        tag_id = raw.replace('/content/cq:tags/', '').strip('/')
        if tag_id in harvest['tags']:
            harvest['tags'][tag_id]['asset_count'] += 1

    print(
        f"   Harvested: "
        f"{len(harvest['nodes'])} nodes, "
        f"{len(harvest['tags'])} tags, "
        f"{len(harvest['namespaces'])} namespaces, "
        f"{len(harvest['folders'])} folders"
    )

    return harvest


def _extract_folder_path(jcr_path: str) -> str:
    """
    Strip /jcr:content and everything below it.
    /content/securian/en/home/jcr:content → /content/securian/en/home
    """
    if '/jcr:content' in jcr_path:
        return jcr_path.split('/jcr:content')[0]
    return jcr_path
