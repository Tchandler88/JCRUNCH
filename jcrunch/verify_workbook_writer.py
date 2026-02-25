import sys, os, shutil
sys.path.insert(0, os.path.join(os.getcwd(), 'jcrunch'))
from export.workbook_writer import write_all_phases
import openpyxl

# Path to your real workbook — update this to match your path
WORKBOOK_PATH = 'AEM_Migration_Analysis_Tool_v3.xlsx'
TEST_OUTPUT   = 'AEM_Migration_Analysis_Tool_v3_TEST.xlsx'

# Safety check — always work on a copy during testing
shutil.copy2(WORKBOOK_PATH, TEST_OUTPUT)
print(f"✓ Working copy created: {TEST_OUTPUT}")

# Build a minimal harvest with one row per phase
harvest = {
    'tags': {
        'wknd-shared/activity/cycling': {
            'tag_id':        'wknd-shared/activity/cycling',
            'tag_title':     'Cycling',
            'parent_tag':    'wknd-shared/activity',
            'depth_level':   3,
            'asset_count':   42,
            'status':        'KEEP - Standard',
            'cloud_notes':   'Map to ACS taxonomy',
            'recommended_map': '/content/cq:tags/wknd-shared/activity/cycling',
            'l1_id':         'wknd-shared',
            'l1_title':      'WKND Shared',
            'l1_desc':       'Root namespace',
            'l2_id':         'wknd-shared/activity',
            'l2_title':      'Activity',
            'l2_desc':       'Activity tags',
            'l3_id':         'wknd-shared/activity/cycling',
            'l3_title':      'Cycling',
            'l3_desc':       'Cycling content',
            'l4_id':         '',
            'l4_title':      '',
            'l4_desc':       '',
            'full_tag_path': '/content/cq:tags/wknd-shared/activity/cycling',
            'tag_label':     'cycling',
        }
    },
    'metadata_fields': {
        'jcr:title': {
            'field_name':           'jcr:title',
            'data_type':            'String',
            'namespace':            'jcr',
            'current_usage_count':  156,
        }
    },
    'workflows': [
        {
            'step_number':    1,
            'step_name':      'DAM Update Asset',
            'step_type':      'Process',
            'fields_affected': 'jcr:title, dc:description',
            'tags_used':       'properties:orientation',
            'conditions':      'mimetype=image/*',
        }
    ],
    'folders': {
        '/content/securian/en': {
            'folder_path':    '/content/securian/en',
            'folder_name':    'en',
            'depth_level':    3,
            'parent_folder':  '/content/securian',
            'child_count':    12,
            'asset_count':    89,
            'is_metadata_like': 'No',
        }
    },
    'namespaces': {
        'http://www.jcp.org/jcr/1.0': {
            'uri':                  'http://www.jcp.org/jcr/1.0',
            'prefix':               'jcr',
            'namespace_id':         'jcr',
            'namespace_type':       'System',
            'used_in':              '156 fields',
            'cloud_support':        'Native Core (Restricted)',
            'fields_in_namespace':  'title, description, primaryType',
            'migration_strategy':   'Lift & Shift',
            'effort':               'Low',
            'timeline_days':        0.5,
        }
    },
}

# Write to test copy
write_all_phases(harvest, TEST_OUTPUT)

# Verify by reading back
print("\n=== VERIFICATION — reading back from workbook ===")
wb = openpyxl.load_workbook(TEST_OUTPUT, data_only=True)

checks = []

# Phase 1 check
ws1 = wb['Phase 1 — Taxonomy Audit']
p1_a4 = ws1['A4'].value
p1_b4 = ws1['B4'].value
p1_f4 = ws1['F4'].value
print(f"Phase 1 A4 (tag_id):    {p1_a4}")
print(f"Phase 1 B4 (tag_title): {p1_b4}")
print(f"Phase 1 F4 (asset_cnt): {p1_f4}")
checks.append(p1_a4 == 'wknd-shared/activity/cycling')
checks.append(p1_b4 == 'Cycling')
checks.append(p1_f4 == 42)

# Phase 2 check
ws2 = wb['Phase 2 — Metadata Schema']
p2_a4 = ws2['A4'].value
p2_h4 = ws2['H4'].value
print(f"\nPhase 2 A4 (field):     {p2_a4}")
print(f"Phase 2 H4 (usage):     {p2_h4}")
checks.append(p2_a4 == 'jcr:title')
checks.append(p2_h4 == 156)

# Phase 3 check
ws3 = wb['Phase 3 — Workflow Extraction']
p3_b4 = ws3['B4'].value
print(f"\nPhase 3 B4 (step_name): {p3_b4}")
checks.append(p3_b4 == 'DAM Update Asset')

# Phase 4 check
ws4 = wb['Phase 4 — Folder Redesign']
p4_a4 = ws4['A4'].value
print(f"\nPhase 4 A4 (folder):    {p4_a4}")
checks.append(p4_a4 == '/content/securian/en')

# Phase 5 check
ws5 = wb['Phase 5 — Namespace Validation']
p5_a4 = ws5['A4'].value
p5_b4 = ws5['B4'].value
print(f"\nPhase 5 A4 (uri):       {p5_a4}")
print(f"Phase 5 B4 (prefix):    {p5_b4}")
checks.append(p5_a4 == 'http://www.jcp.org/jcr/1.0')
checks.append(p5_b4 == 'jcr')

print(f"\nAll checks: {'PASS' if all(checks) else 'FAIL'}")
if not all(checks):
    for i, c in enumerate(checks):
        if not c:
            print(f"  ✗ Check {i+1} failed")

overall = all(checks)
print(f"\n{'✓ STEP 5 COMPLETE — workbook writer is working' if overall else '✗ STEP 5 FAILED — review output above'}")
print(f"{'  Open AEM_Migration_Analysis_Tool_v3_TEST.xlsx to visually verify' if overall else ''}")
