import re


# Namespaces that are system-managed in AEM
SYSTEM_NAMESPACES = {
    'jcr', 'oak', 'sling', 'granite',
    'rep', 'nt', 'mix', 'vlt', 'cq'
}


def run_metadata_audit(harvest: dict):
    """
    Aggregates harvest['properties'] into harvest['metadata_fields'].
    Each unique full_name becomes one metadata field row.
    Computes: data_type, is_system_managed, usage_count, anomaly_flags.
    No database. No file writes. Mutates harvest dict only.
    """
    properties = harvest.get('properties', {})
    if not properties:
        print("   [!] No properties found in harvest — skipping")
        harvest['metadata_fields'] = {}
        return

    # Aggregate — one entry per unique full_name
    # Track: node paths that use it, all values seen
    aggregated = {}  # {full_name: {namespace, name, node_paths, values}}

    for prop in properties.values():
        full_name  = prop.get('full_name', '')
        namespace  = prop.get('namespace', '')
        name       = prop.get('name', '')
        value      = prop.get('value', '')
        jcr_path   = prop.get('jcr_path', '')

        if not full_name:
            continue

        if full_name not in aggregated:
            aggregated[full_name] = {
                'full_name':  full_name,
                'namespace':  namespace,
                'name':       name,
                'node_paths': set(),
                'values':     [],
            }

        aggregated[full_name]['node_paths'].add(jcr_path)
        if value:
            aggregated[full_name]['values'].append(value)

    # Build metadata_fields dict
    metadata_fields = {}

    for full_name, agg in aggregated.items():
        namespace   = agg['namespace']
        usage_count = len(agg['node_paths'])

        # Infer data type from sampled values
        sample_values = agg['values']
        data_type = _infer_data_type(sample_values)

        # System managed flag
        is_system = 'Yes' if namespace in SYSTEM_NAMESPACES else 'No'

        # Anomaly flags
        flags = []
        if usage_count == 0:
            flags.append('UNUSED - consider deprecation')
        if is_system == 'No':
            flags.append('No cloud equivalent mapped')
        anomaly_flags = ' | '.join(flags)

        metadata_fields[full_name] = {
            'field_name':           full_name,
            'namespace':            namespace,
            'data_type':            data_type,
            'is_system_managed':    is_system,
            'current_usage_count':  usage_count,
            'anomaly_flags':        anomaly_flags,
        }

    harvest['metadata_fields'] = metadata_fields
    print(f"   [ok] Metadata audit complete: "
          f"{len(metadata_fields)} unique fields aggregated")


def _infer_data_type(values: list) -> str:
    """
    Infer data type from a list of observed values.
    Checks the first non-empty value found.
    First match wins.
    """
    for value in values:
        if not value:
            continue
        v = str(value).strip()
        if not v:
            continue

        if v.lower() in ('true', 'false'):
            return 'Boolean'
        if v.isdigit():
            return 'Long'
        if re.match(r'^\d{4}-\d{2}-\d{2}', v):
            return 'Date'
        if v.startswith('/content/'):
            return 'Path Reference'
        return 'String'

    return 'String'
