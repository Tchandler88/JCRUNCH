def run_namespace_audit(harvest: dict):
    """
    Enriches harvest['namespaces'] in place.
    Adds all derived columns needed by Phase 5 workbook sheet.
    No database. No file writes. Mutates harvest dict only.
    """
    namespaces = harvest.get('namespaces', {})
    if not namespaces:
        print("   [!] No namespaces found in harvest — skipping")
        return

    # Pre-compute property usage counts per namespace prefix
    # harvest['properties'] is a dict keyed by (jcr_path, full_name)
    prefix_field_counts  = {}   # {prefix: int}
    prefix_field_names   = {}   # {prefix: set of field names}

    for prop in harvest.get('properties', {}).values():
        prefix = prop.get('namespace', '')
        name   = prop.get('name', '')
        if prefix:
            prefix_field_counts[prefix] = \
                prefix_field_counts.get(prefix, 0) + 1
            if prefix not in prefix_field_names:
                prefix_field_names[prefix] = set()
            prefix_field_names[prefix].add(name)

    enriched = 0
    for uri, ns in namespaces.items():
        prefix = ns.get('prefix', '')

        ns_type  = _classify_type(uri)
        support  = _classify_support(ns_type)
        strategy = _classify_strategy(support)
        effort   = _classify_effort(strategy)
        timeline = _classify_timeline(effort)

        # used_in — count of properties using this prefix
        count   = prefix_field_counts.get(prefix, 0)
        used_in = f"{count} fields"

        # fields_in_namespace — unique field names, max 10
        names = sorted(prefix_field_names.get(prefix, set()))
        if len(names) > 10:
            fields_str = ', '.join(names[:10]) + '...'
        else:
            fields_str = ', '.join(names)

        ns.update({
            'namespace_id':        prefix,
            'namespace_type':      ns_type,
            'cloud_support':       support,
            'migration_strategy':  strategy,
            'effort':              effort,
            'timeline_days':       timeline,
            'used_in':             used_in,
            'fields_in_namespace': fields_str,
        })
        enriched += 1

    print(f"   [ok] Namespace audit complete: {enriched} namespaces enriched")


def _classify_type(uri: str) -> str:
    uri_lower = uri.lower()
    if any(t in uri_lower for t in
           ['jcr', 'oak', 'sling', 'granite', 'day.com/jcr']):
        return 'System (Repo)'
    if any(t in uri_lower for t in
           ['w3.org', 'purl.org', 'iptc', 'cipa', 'prism']):
        return 'Standard'
    if 'adobe' in uri_lower:
        return 'Vendor (Adobe)'
    if 'microsoft' in uri_lower:
        return 'Vendor (Microsoft)'
    if 'apple' in uri_lower:
        return 'Vendor (Apple)'
    return 'Custom'


def _classify_support(namespace_type: str) -> str:
    if namespace_type == 'System (Repo)':
        return 'Native Core (Restricted)'
    if namespace_type == 'Standard':
        return 'Native Supported'
    return 'Requires CND'


def _classify_strategy(cloud_support: str) -> str:
    return {
        'Native Core (Restricted)': 'Do Not Migrate',
        'Native Supported':         'Lift & Shift',
        'Requires CND':             'Register CND',
    }.get(cloud_support, 'Manual Review')


def _classify_effort(migration_strategy: str) -> str:
    return {
        'Do Not Migrate': 'Low',
        'Lift & Shift':   'Low',
        'Register CND':   'Medium',
        'Manual Review':  'High',
    }.get(migration_strategy, 'High')


def _classify_timeline(effort: str) -> float:
    return {
        'Low':    0.5,
        'Medium': 3.0,
        'High':   7.0,
    }.get(effort, 7.0)
