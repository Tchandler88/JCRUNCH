import re


METADATA_LIKE_ORIENTATIONS = {'landscape', 'portrait', 'square'}
METADATA_LIKE_STATES       = {'approved', 'archive', 'archived'}
METADATA_LIKE_COLORS       = {'color', 'colour'}
METADATA_LIKE_REGIONS      = {
    'north', 'south', 'east', 'west',
    'apac', 'emea', 'nam', 'latam'
}


def run_folder_audit(harvest: dict):
    """
    Enriches harvest['folders'] in place.
    Adds child_count, asset_count, is_metadata_like to every folder.
    No database. No file writes. Mutates harvest dict only.
    """
    folders = harvest.get('folders', {})
    if not folders:
        print("   [!] No folders found in harvest — skipping")
        return

    nodes = harvest.get('nodes', {})

    # Pre-compute child counts
    # Count how many folders list each path as their parent_folder
    child_counts = {}
    for f in folders.values():
        parent = f.get('parent_folder', '')
        if parent:
            child_counts[parent] = child_counts.get(parent, 0) + 1

    # Pre-compute asset counts per folder prefix
    # dam:Asset nodes only
    asset_counts = {}
    for node in nodes.values():
        if node.get('node_type') == 'dam:Asset':
            node_path = node.get('path', '')
            for folder_path in folders:
                prefix = folder_path.rstrip('/') + '/'
                if node_path.startswith(prefix):
                    asset_counts[folder_path] = \
                        asset_counts.get(folder_path, 0) + 1

    enriched = 0
    for folder_path, folder in folders.items():
        folder_name = folder.get('folder_name', '')

        folder.update({
            'child_count':      child_counts.get(folder_path, 0),
            'asset_count':      asset_counts.get(folder_path, 0),
            'is_metadata_like': _is_metadata_like(folder_name),
        })
        enriched += 1

    print(f"   [ok] Folder audit complete: {enriched} folders enriched")


def _is_metadata_like(name: str) -> str:
    """
    Returns 'Yes' if folder name matches a metadata-like pattern.
    Returns 'No' otherwise.
    """
    if not name:
        return 'No'

    n = name.strip().lower()

    if re.match(r'^\d+$', n):
        return 'Yes'
    if re.match(r'^q[1-4]$', n):
        return 'Yes'
    if n in METADATA_LIKE_ORIENTATIONS:
        return 'Yes'
    if n in METADATA_LIKE_STATES:
        return 'Yes'
    if n in METADATA_LIKE_COLORS:
        return 'Yes'
    if n in METADATA_LIKE_REGIONS:
        return 'Yes'

    return 'No'
