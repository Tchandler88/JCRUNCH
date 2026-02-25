def build_tag_hierarchy(tag_id: str, tag_lookup: dict) -> dict:
    """
    Given a tag ID like 'wknd-shared/activity/cycling',
    build the L1-L4 hierarchy columns.

    tag_lookup = {tag_id: {'tag_title': ..., 'description': ...}}

    Returns dict with l1_id, l1_title, l1_desc ... l4_id, l4_title, l4_desc
    """
    parts = tag_id.split('/')
    result = {}
    for level in range(1, 5):
        if level <= len(parts):
            ancestor_id = '/'.join(parts[:level])
            ancestor    = tag_lookup.get(ancestor_id, {})
            result[f'l{level}_id']    = ancestor_id
            result[f'l{level}_title'] = ancestor.get('tag_title', '')
            result[f'l{level}_desc']  = ancestor.get('description', '')
        else:
            result[f'l{level}_id']    = ''
            result[f'l{level}_title'] = ''
            result[f'l{level}_desc']  = ''
    return result


def calculate_depth(tag_id: str) -> int:
    return tag_id.count('/') + 1


def extract_parent(tag_id: str) -> str:
    if '/' not in tag_id:
        return ''
    return tag_id.rsplit('/', 1)[0]


def extract_label(tag_id: str) -> str:
    return tag_id.rsplit('/', 1)[-1]
