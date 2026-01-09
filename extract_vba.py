"""
MS Access VBA Code Extractor
Extracts all VBA code from an Access database including modules, classes, forms, and reports
"""

import win32com.client
import os
import sys
from pathlib import Path

def extract_vba_from_access(accdb_path, output_dir):
    """
    Extract all VBA code from an MS Access database file.
    
    Args:
        accdb_path: Path to the .accdb or .mdb file
        output_dir: Directory where VBA files will be saved
    """
    
    print(f"Opening Access database: {accdb_path}")
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Initialize Access Application
    access = win32com.client.Dispatch("Access.Application")
    
    try:
        # Open the database (convert to absolute path)
        abs_path = os.path.abspath(accdb_path)
        access.OpenCurrentDatabase(abs_path)
        
        # Get VBA project
        vba_project = access.VBE.ActiveVBProject
        
        print(f"\nProject Name: {vba_project.Name}")
        print(f"Total Components: {vba_project.VBComponents.Count}\n")
        
        # Track statistics
        stats = {
            'modules': 0,
            'class_modules': 0,
            'forms': 0,
            'reports': 0,
            'total_lines': 0,
            'components': []
        }
        
        # Iterate through all VBA components
        for component in vba_project.VBComponents:
            component_name = component.Name
            component_type = component.Type
            
            # Determine component type
            type_name = {
                1: "module",           # vbext_ct_StdModule
                2: "class",            # vbext_ct_ClassModule
                3: "form",             # vbext_ct_MSForm
                100: "document"        # vbext_ct_Document (Form/Report)
            }.get(component_type, "unknown")
            
            print(f"Extracting: {component_name} ({type_name})")
            
            # Get code from component
            try:
                code_module = component.CodeModule
                line_count = code_module.CountOfLines
                
                if line_count > 0:
                    code_text = code_module.Lines(1, line_count)
                    
                    # Determine file extension
                    if component_type == 1:  # Standard module
                        ext = "bas"
                        stats['modules'] += 1
                    elif component_type == 2:  # Class module
                        ext = "cls"
                        stats['class_modules'] += 1
                    elif component_type in [3, 100]:  # Form or Report
                        if "Form_" in component_name or component_type == 3:
                            ext = "form.bas"
                            stats['forms'] += 1
                        else:
                            ext = "report.bas"
                            stats['reports'] += 1
                    else:
                        ext = "vba"
                    
                    # Save to file
                    output_file = os.path.join(output_dir, f"{component_name}.{ext}")
                    with open(output_file, 'w', encoding='utf-8', errors='replace') as f:
                        f.write(f"' Component: {component_name}\n")
                        f.write(f"' Type: {type_name}\n")
                        f.write(f"' Lines: {line_count}\n")
                        f.write("' " + "="*60 + "\n\n")
                        f.write(code_text)
                    
                    stats['total_lines'] += line_count
                    stats['components'].append({
                        'name': component_name,
                        'type': type_name,
                        'lines': line_count,
                        'file': output_file
                    })
                    print(f"  [OK] Saved {line_count} lines to {os.path.basename(output_file)}")
                else:
                    print(f"  - No code found")
                    
            except Exception as e:
                print(f"  [ERROR] Error extracting code: {e}")
        
        # Print summary
        print("\n" + "="*70)
        print("EXTRACTION SUMMARY")
        print("="*70)
        print(f"Standard Modules: {stats['modules']}")
        print(f"Class Modules: {stats['class_modules']}")
        print(f"Forms with Code: {stats['forms']}")
        print(f"Reports with Code: {stats['reports']}")
        print(f"Total Lines of Code: {stats['total_lines']}")
        print(f"\nAll files saved to: {output_dir}")
        
        return stats
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        # Close Access
        try:
            access.CloseCurrentDatabase()
            access.Quit()
            print("\nAccess closed.")
        except:
            pass


if __name__ == "__main__":
    # Configuration
    script_dir = Path(__file__).parent
    accdb_path = str(script_dir / "msaccess" / "TB CMS.SQL.accdb")
    output_dir = str(script_dir / "extracted_vba")
    
    if not os.path.exists(accdb_path):
        print(f"Error: Access database not found at {accdb_path}")
        sys.exit(1)
    
    print("="*70)
    print("MS Access VBA Code Extractor")
    print("="*70)
    print()
    
    result = extract_vba_from_access(accdb_path, output_dir)
    
    if result:
        print("\n[SUCCESS] Extraction completed successfully!")
    else:
        print("\n[FAILED] Extraction failed. Please check errors above.")
        sys.exit(1)
