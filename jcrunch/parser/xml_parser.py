import xml.etree.ElementTree as ET
import re

MULTI_VALUE_PATTERN = re.compile(r'^\[(.+)\]$')

def parse_content_xml(xml_path: str, jcr_path: str) -> dict:
    """
    Parse a single AEM .content.xml file.

    Input XML example:
      <jcr:root xmlns:jcr="http://www.jcp.org/jcr/1.0"
                xmlns:cq="http://www.day.com/jcr/cq/1.0"
          jcr:primaryType="cq:Page"
          jcr:title="Home"
          cq:tags="[wknd-shared/activity/cycling,properties:orientation/landscape]"
          cq:template="/conf/securian/settings/wcm/templates/homepage"/>

    Output dict:
      {
        'path': '/content/securian/en',
        'node_type': 'cq:Page',
        'resource_type': None,
        'template': '/conf/securian/settings/wcm/templates/homepage',
        'last_modified': None,
        'last_modified_by': None,
        'namespaces': {'jcr': 'http://www.jcp.org/jcr/1.0', ...},
        'properties': [
            {'namespace': 'jcr', 'name': 'title',
             'full_name': 'jcr:title', 'value': 'Home', 'is_multi': False}
        ],
        'tags': ['wknd-shared/activity/cycling',
                 'properties:orientation/landscape']
      }
    """
    # ElementTree strips xmlns: declarations from root.attrib.
    # Use iterparse with start-ns to capture them before they disappear.
    namespaces = {}
    root = None
    try:
        for event, elem in ET.iterparse(xml_path, events=['start-ns', 'start']):
            if event == 'start-ns':
                prefix, uri = elem
                namespaces[prefix] = uri
            elif event == 'start':
                root = elem
                break
    except ET.ParseError:
        with open(xml_path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        root = ET.fromstring(content)
        # Re-extract any xmlns: that survived as plain attribs (fallback only)
        for key, val in root.attrib.items():
            if key.startswith('xmlns:'):
                namespaces[key[6:]] = val

    result = {
        'path': jcr_path,
        'node_type': None,
        'resource_type': None,
        'template': None,
        'last_modified': None,
        'last_modified_by': None,
        'namespaces': namespaces,
        'properties': [],
        'tags': []
    }

    for attr_key, attr_val in root.attrib.items():
        if attr_key.startswith('xmlns:') or attr_key == 'xmlns':
            continue

        prop_name = _clark_to_prefixed(attr_key, namespaces)
        namespace  = prop_name.split(':')[0] if ':' in prop_name else ''
        local_name = prop_name.split(':')[1] if ':' in prop_name else prop_name

        # Top-level node fields — these go into Node table directly
        if prop_name == 'jcr:primaryType':
            result['node_type'] = attr_val
            continue
        if prop_name == 'sling:resourceType':
            result['resource_type'] = attr_val
            continue
        if prop_name == 'cq:template':
            result['template'] = attr_val
            continue
        if prop_name in ('jcr:lastModified', 'cq:lastModified'):
            result['last_modified'] = attr_val
            continue
        if prop_name in ('jcr:lastModifiedBy', 'cq:lastModifiedBy'):
            result['last_modified_by'] = attr_val
            continue

        # Multi-value properties: [val1,val2,val3]
        is_multi = False
        values = [attr_val]
        mv_match = MULTI_VALUE_PATTERN.match(attr_val)
        if mv_match:
            is_multi = True
            values = _split_multivalue(mv_match.group(1))

        # Tag assignments
        if prop_name in ('cq:tags', 'sling:tags'):
            result['tags'].extend(values)
            continue

        for v in values:
            result['properties'].append({
                'namespace': namespace,
                'name': local_name,
                'full_name': prop_name,
                'value': v,
                'is_multi': is_multi
            })

    return result


def _clark_to_prefixed(attr_key: str, namespaces: dict) -> str:
    """Convert {uri}localname to prefix:localname."""
    if attr_key.startswith('{'):
        uri_end = attr_key.index('}')
        uri   = attr_key[1:uri_end]
        local = attr_key[uri_end+1:]
        prefix = next((k for k, v in namespaces.items() if v == uri), None)
        return f"{prefix}:{local}" if prefix else local
    return attr_key


def _split_multivalue(raw: str) -> list:
    """
    Split AEM multi-value string. Handles escaped commas.
    '[wknd-shared/activity/cycling,properties:orientation/landscape]'
    → ['wknd-shared/activity/cycling', 'properties:orientation/landscape']
    """
    return [v.strip() for v in raw.split(',') if v.strip()]
