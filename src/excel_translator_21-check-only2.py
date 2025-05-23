#!/usr/bin/env python3
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

def diagnose_excel_file(file_path):
    print(f"Analyzing: {file_path}")
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_file:
            files = zip_file.namelist()
            print(f"✅ Valid ZIP with {len(files)} internal files")
            
            # Check critical files
            critical = ['xl/workbook.xml', 'xl/styles.xml', '[Content_Types].xml']
            for crit in critical:
                if crit in files:
                    print(f"✅ {crit} present")
                else:
                    print(f"❌ {crit} missing")
            
            # Analyze styles.xml
            if 'xl/styles.xml' in files:
                styles_content = zip_file.read('xl/styles.xml')
                print(f"Styles file size: {len(styles_content)} bytes")
                
                try:
                    root = ET.fromstring(styles_content)
                    
                    # Count style elements
                    cellXfs = root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cellXfs')
                    cellStyleXfs = root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cellStyleXfs')
                    
                    cellXfs_count = len(cellXfs) if cellXfs is not None else 0
                    cellStyleXfs_count = len(cellStyleXfs) if cellStyleXfs is not None else 0
                    
                    print(f"cellXfs count: {cellXfs_count}")
                    print(f"cellStyleXfs count: {cellStyleXfs_count}")
                    
                    # Check for named styles
                    named_styles = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cellStyle')
                    print(f"Named styles count: {len(named_styles)}")
                    
                    # Look for problematic references
                    for ns in named_styles:
                        xf_id = ns.get('xfId')
                        if xf_id:
                            xf_id_int = int(xf_id)
                            if xf_id_int >= cellStyleXfs_count:
                                print(f"❌ PROBLEM: Named style references xfId {xf_id_int} but only {cellStyleXfs_count} cellStyleXfs exist")
                    
                except Exception as e:
                    print(f"❌ Styles XML parsing error: {e}")
            
            # Get sheet info
            if 'xl/workbook.xml' in files:
                wb_content = zip_file.read('xl/workbook.xml')
                wb_root = ET.fromstring(wb_content)
                sheets = wb_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')
                print(f"Sheets found: {len(sheets)}")
                for sheet in sheets:
                    print(f"  - {sheet.get('name')}")
    
    except Exception as e:
        print(f"❌ Analysis failed: {e}")

# Run the diagnosis
if __name__ == "__main__":
    file_path = "../input_files/Draft.xlsx"
    diagnose_excel_file(file_path)