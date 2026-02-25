# JCRUNCH — JCR Content Repository Unifier and Node-to-Column Harvester
# It's GR-R-REAT for metadata audits.

import click
import os
import sys

# Ensure imports resolve correctly when called from VBA (working dir may differ)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


@click.command()
@click.option('--package',
    type=click.Path(exists=True),
    help='AEM Package Manager .zip file')
@click.option('--workbook',
    type=click.Path(),
    required=True,
    help='Path to AEM_Migration_Analysis_Tool_v3.xlsx')
@click.option('--run-ai',
    is_flag=True, default=False,
    help='Run AI Bot after parsing')
@click.option('--ai-only',
    is_flag=True, default=False,
    help='Skip parsing, only run AI fills on existing workbook')
@click.option('--phase',
    default='all',
    help='Run specific phase: 1,2,3,4,5 or all')
def main(package, workbook, run_ai, ai_only, phase):

    print("JCRUNCH -- It's GR-R-REAT for metadata audits")

    harvest = {}

    if not ai_only and package:
        from parser.package_reader import walk_package
        from audit.tag_auditor import run_tag_audit
        from audit.namespace_auditor import run_namespace_audit
        from audit.metadata_auditor import run_metadata_audit
        from audit.folder_auditor import run_folder_audit

        print(f"Reading package: {package}")
        harvest = walk_package(package)

        if phase in ('all', '1'):
            run_tag_audit(harvest)
            print("   Phase 1 tag audit complete")
        if phase in ('all', '2'):
            run_metadata_audit(harvest)
            print("   Phase 2 metadata audit complete")
        if phase in ('all', '4'):
            run_folder_audit(harvest)
            print("   Phase 4 folder audit complete")
        if phase in ('all', '5'):
            run_namespace_audit(harvest)
            print("   Phase 5 namespace audit complete")

    if run_ai:
        from ai.bot import run_ai_fills
        print("Running AI Bot fills...")
        run_ai_fills(harvest, workbook, phase=phase)
        print("   AI fills complete")

    print(f"Writing to workbook: {workbook}")
    from export.workbook_writer import write_all_phases
    write_all_phases(harvest, workbook)
    print("   Workbook populated")
    print("JCRUNCH done. Open your workbook.")


if __name__ == '__main__':
    main()
