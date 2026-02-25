"""
inject_ribbon.py — JCRUNCH Ribbon Injector

Injects the JCRUNCH custom ribbon tab into an Excel workbook (.xlsx or .xlsm)
without manual zip editing, which is unreliable on Windows.

Usage:
    python vba/inject_ribbon.py --workbook "AEM_Migration_Analysis_Tool_v3.xlsx"

Output:
    AEM_Migration_Analysis_Tool_v3_ribbon.xlsm  (written next to the original)

The original workbook is never modified. Re-running is safe.
"""

import argparse
import io
import os
import re
import sys
import zipfile

# ─────────────────────────────────────────────────────────────────
# RIBBON XML — embedded verbatim so the script is self-contained
# ─────────────────────────────────────────────────────────────────
RIBBON_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabJCRUNCH"
           label="JCRUNCH"
           insertAfterMso="TabView">
        <group id="grpPackage"
               label="AEM Package">
          <button id="btnBrowse"
                  label="Browse Package"
                  screentip="Select AEM Package .zip"
                  supertip="Open a file dialog to choose the AEM Package Manager export (.zip) to analyze."
                  size="large"
                  imageMso="FolderOpen"
                  onAction="OnBrowsePackage"/>
          <button id="btnRun"
                  label="Run JCRUNCH"
                  screentip="Run the JCRUNCH pipeline"
                  supertip="Parse the selected AEM package and populate all Phase sheets in this workbook."
                  size="large"
                  imageMso="PlaybackPlay"
                  onAction="OnRunJCRUNCH"/>
        </group>
        <group id="grpStatus"
               label="Status">
          <labelControl id="lblStatus"
                        label="Select a package to begin"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""

CUSTOMUI_PATH = "customUI/customUI14.xml"
REL_TYPE      = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
REL_ID        = "rId_jcrunch"


def patch_rels(original_xml: bytes) -> bytes:
    """
    Add the customUI relationship to _rels/.rels if not already present.
    Uses simple string manipulation to avoid re-serialising the XML
    (which would strip namespace declarations and break the file).
    """
    text = original_xml.decode("utf-8")

    # Idempotent: skip if already injected
    if CUSTOMUI_PATH in text:
        print("  [i] _rels/.rels already contains customUI entry — skipping")
        return original_xml

    new_rel = (
        f'<Relationship Id="{REL_ID}" '
        f'Type="{REL_TYPE}" '
        f'Target="{CUSTOMUI_PATH}"/>'
    )

    # Insert just before the closing </Relationships> tag
    if "</Relationships>" in text:
        text = text.replace("</Relationships>", f"  {new_rel}\n</Relationships>")
    else:
        raise ValueError("_rels/.rels has unexpected format — </Relationships> not found")

    return text.encode("utf-8")


def patch_content_types(original_xml: bytes) -> bytes:
    """
    Add the customUI14.xml Override to [Content_Types].xml if not already present.
    Also uses string manipulation for the same safety reason as patch_rels.
    """
    text = original_xml.decode("utf-8")

    # Idempotent
    if "customUI14.xml" in text:
        print("  [i] [Content_Types].xml already contains customUI entry — skipping")
        return original_xml

    # If there's already a Default for .xml extension we don't need an Override.
    # But an explicit Override never hurts and is always safe to add.
    new_override = '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>'

    if "</Types>" in text:
        text = text.replace("</Types>", f"  {new_override}\n</Types>")
    else:
        raise ValueError("[Content_Types].xml has unexpected format — </Types> not found")

    return text.encode("utf-8")


def inject(workbook_path: str) -> str:
    workbook_path = os.path.abspath(workbook_path)

    if not os.path.exists(workbook_path):
        print(f"[!] Workbook not found: {workbook_path}")
        sys.exit(1)

    # Build output path — always .xlsm (macro-enabled)
    stem = re.sub(r"\.(xlsx|xlsm)$", "", os.path.basename(workbook_path), flags=re.IGNORECASE)
    out_name = stem + "_ribbon.xlsm"
    out_path = os.path.join(os.path.dirname(workbook_path), out_name)

    print(f"[>>] Reading:  {workbook_path}")
    print(f"[>>] Writing:  {out_path}")

    # ── Read all entries from the original zip ─────────────────────
    entries: dict[str, bytes] = {}
    with zipfile.ZipFile(workbook_path, "r") as zin:
        for name in zin.namelist():
            entries[name] = zin.read(name)

    # ── Patch _rels/.rels ──────────────────────────────────────────
    if "_rels/.rels" not in entries:
        print("[!] _rels/.rels not found in workbook — is this a valid .xlsx/.xlsm?")
        sys.exit(1)

    print("  [>>] Patching _rels/.rels ...")
    entries["_rels/.rels"] = patch_rels(entries["_rels/.rels"])

    # ── Patch [Content_Types].xml ──────────────────────────────────
    if "[Content_Types].xml" not in entries:
        print("[!] [Content_Types].xml not found — is this a valid .xlsx/.xlsm?")
        sys.exit(1)

    print("  [>>] Patching [Content_Types].xml ...")
    entries["[Content_Types].xml"] = patch_content_types(entries["[Content_Types].xml"])

    # ── Add the ribbon XML ─────────────────────────────────────────
    if CUSTOMUI_PATH in entries:
        print(f"  [i] {CUSTOMUI_PATH} already exists — overwriting with current ribbon XML")
    else:
        print(f"  [>>] Adding {CUSTOMUI_PATH} ...")
    entries[CUSTOMUI_PATH] = RIBBON_XML.encode("utf-8")

    # ── Write new zip ──────────────────────────────────────────────
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    with open(out_path, "wb") as f:
        f.write(buf.getvalue())

    print(f"\n[ok] Done!  Output: {out_path}")
    print()
    print("Next steps:")
    print("  1. Open the _ribbon.xlsm file in Excel")
    print("  2. If prompted, click 'Enable Content' to allow macros")
    print("  3. The JCRUNCH tab should appear in the ribbon")
    print("  4. If the VBA module is not yet imported, press Alt+F11,")
    print("     File > Import File, and select jcrunch/vba/JCRUNCH_Ribbon.bas")

    return out_path


def main():
    parser = argparse.ArgumentParser(
        description="Inject the JCRUNCH ribbon tab into an Excel workbook."
    )
    parser.add_argument(
        "--workbook", required=True,
        help="Path to AEM_Migration_Analysis_Tool_v3.xlsx (or any .xlsx/.xlsm)"
    )
    args = parser.parse_args()
    inject(args.workbook)


if __name__ == "__main__":
    main()
