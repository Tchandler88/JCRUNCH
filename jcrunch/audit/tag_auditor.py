import re
from parser.tag_resolver import (
    build_tag_hierarchy,
    calculate_depth,
    extract_parent,
    extract_label,
)


def run_tag_audit(harvest: dict):
    """
    Enriches harvest['tags'] in place.
    Adds all derived columns needed by Phase 1 workbook sheet.
    No database. No file writes. Mutates harvest dict only.
    """
    tags = harvest.get('tags', {})
    if not tags:
        print("   [!] No tags found in harvest — skipping tag audit")
        return

    # Build lookup dict for hierarchy resolution
    # {tag_id: {'tag_title': ..., 'description': ...}}
    tag_lookup = {
        tag_id: {
            'tag_title':   t.get('tag_title', ''),
            'description': t.get('description', ''),
        }
        for tag_id, t in tags.items()
    }

    # Build title frequency map for duplicate detection
    # Must be built BEFORE the loop
    title_counts = {}
    for t in tags.values():
        title = (t.get('tag_title') or '').strip()
        if title:
            title_counts[title] = title_counts.get(title, 0) + 1

    enriched = 0
    for tag_id, tag in tags.items():

        depth   = calculate_depth(tag_id)
        parent  = extract_parent(tag_id)
        label   = extract_label(tag_id)

        title       = (tag.get('tag_title') or '').strip()
        asset_count = tag.get('asset_count', 0)

        status      = _calculate_status(
            tag_id, title, asset_count,
            title_counts, depth
        )
        cloud_notes = _calculate_cloud_notes(status)
        rec_map     = f"/content/cq:tags/{tag_id}"
        full_path   = f"/content/cq:tags/{tag_id}"

        hierarchy   = build_tag_hierarchy(tag_id, tag_lookup)

        # Mutate the tag dict in place — add all derived keys
        tag.update({
            'depth_level':    depth,
            'parent_tag':     parent,
            'tag_label':      label,
            'status':         status,
            'cloud_notes':    cloud_notes,
            'recommended_map': rec_map,
            'full_tag_path':  full_path,
            'l1_id':          hierarchy.get('l1_id', ''),
            'l1_title':       hierarchy.get('l1_title', ''),
            'l1_desc':        hierarchy.get('l1_desc', ''),
            'l2_id':          hierarchy.get('l2_id', ''),
            'l2_title':       hierarchy.get('l2_title', ''),
            'l2_desc':        hierarchy.get('l2_desc', ''),
            'l3_id':          hierarchy.get('l3_id', ''),
            'l3_title':       hierarchy.get('l3_title', ''),
            'l3_desc':        hierarchy.get('l3_desc', ''),
            'l4_id':          hierarchy.get('l4_id', ''),
            'l4_title':       hierarchy.get('l4_title', ''),
            'l4_desc':        hierarchy.get('l4_desc', ''),
        })
        enriched += 1

    print(f"   [ok] Tag audit complete: {enriched} tags enriched")


def _calculate_status(
    tag_id: str,
    tag_title: str,
    asset_count: int,
    title_counts: dict,
    depth: int
) -> str:
    """
    Gatekeeper priority chain — first match wins.
    Mirrors the Column G formula from the workbook exactly.
    """
    leaf = extract_label(tag_id)

    # Priority 1 — Missing Title
    if not tag_title:
        return 'DEPRECATE - Missing Title'

    # Priority 2 — Bad Naming (uppercase or space in leaf segment)
    if re.search(r'[A-Z\s]', leaf):
        return 'DEPRECATE - Bad Naming'

    # Priority 3 — Obsolete keywords anywhere in tag_id
    if re.search(
        r'\b(test|temp|mock|old|delete|backup|draft)\b',
        tag_id,
        re.IGNORECASE
    ):
        return 'DEPRECATE - Obsolete'

    # Priority 4 — Duplicate Title
    if title_counts.get(tag_title, 0) > 1:
        return 'CONSOLIDATE - Duplicate Title'

    # Priority 5 — Zero Usage
    if asset_count == 0:
        return 'REVIEW - Zero Usage'

    # Priority 6 — High Usage
    if asset_count > 100:
        return 'REVIEW - High Usage'

    # Priority 7 — Too Deep
    if depth > 4:
        return 'REVIEW - Too Deep'

    # Priority 8 — Default
    return 'KEEP - Standard'


def _calculate_cloud_notes(status: str) -> str:
    """
    Translate status into a human-readable migration action.
    Mirrors Column H translation formula from the workbook.
    """
    notes_map = {
        'DEPRECATE - Missing Title':     'Do not migrate — no title defined',
        'DEPRECATE - Bad Naming':        'Do not migrate — fix naming convention first',
        'DEPRECATE - Obsolete':          'Do not migrate — obsolete tag detected',
        'CONSOLIDATE - Duplicate Title': 'Merge with duplicate before migrating',
        'REVIEW - Zero Usage':           'Audit required — tag is unused',
        'REVIEW - High Usage':           'Audit required — high usage, verify mapping',
        'REVIEW - Too Deep':             'Audit required — exceeds recommended depth',
        'KEEP - Standard':               'Migrate as-is',
    }
    return notes_map.get(status, 'Manual review required')
