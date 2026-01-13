"""
MS Access Database Comprehensive Assessment Tool
Extracts complete information about tables, queries, and VBA code from an Access database
"""

import win32com.client
import os
import sys
import time
import json
from pathlib import Path
from datetime import datetime


def extract_table_info(access) -> list[dict]:
    """
    Extract information about all tables in the database.
    
    Args:
        access: Active Access Application COM object
        
    Returns:
        List of dictionaries containing table information
    """
    tables_info = []
    
    try:
        current_db = access.CurrentDb()
        table_defs = current_db.TableDefs
        
        print(f"\nExtracting information from {table_defs.Count} table definitions...")
        
        for table_def in table_defs:
            try:
                table_name = table_def.Name
                
                # Skip system tables (start with MSys or ~)
                if table_name.startswith('MSys') or table_name.startswith('~'):
                    continue
                
                print(f"  - Analyzing table: {table_name}")
                
                # Get fields information
                fields = []
                for field in table_def.Fields:
                    try:
                        field_info = {
                            'name': field.Name,
                            'type': get_field_type_name(field.Type),
                            'type_code': field.Type,
                            'size': field.Size if hasattr(field, 'Size') else None,
                            'required': field.Required if hasattr(field, 'Required') else False,
                            'allow_zero_length': field.AllowZeroLength if hasattr(field, 'AllowZeroLength') else False,
                        }
                        
                        # Try to get default value
                        try:
                            field_info['default_value'] = field.DefaultValue
                        except:
                            field_info['default_value'] = None
                        
                        fields.append(field_info)
                    except Exception as e:
                        print(f"    [WARNING] Error reading field: {e}")
                        continue
                
                # Get indexes information
                indexes = []
                try:
                    for index in table_def.Indexes:
                        try:
                            index_fields = []
                            for idx_field in index.Fields:
                                index_fields.append(idx_field.Name)
                            
                            indexes.append({
                                'name': index.Name,
                                'primary': index.Primary if hasattr(index, 'Primary') else False,
                                'unique': index.Unique if hasattr(index, 'Unique') else False,
                                'fields': index_fields
                            })
                        except Exception as e:
                            print(f"    [WARNING] Error reading index: {e}")
                            continue
                except:
                    pass  # Some tables may not have indexes
                
                # Get record count (if possible)
                record_count = None
                try:
                    recordset = current_db.OpenRecordset(f"SELECT COUNT(*) FROM [{table_name}]")
                    if not recordset.EOF:
                        record_count = recordset.Fields(0).Value
                    recordset.Close()
                except:
                    pass  # Can't get record count for some tables
                
                tables_info.append({
                    'name': table_name,
                    'fields': fields,
                    'indexes': indexes,
                    'record_count': record_count,
                    'field_count': len(fields)
                })
                
            except Exception as e:
                print(f"  [ERROR] Error extracting table {table_name}: {e}")
                continue
        
        print(f"  [OK] Extracted information from {len(tables_info)} tables")
        
    except Exception as e:
        print(f"[ERROR] Error accessing table definitions: {e}")
        import traceback
        traceback.print_exc()
    
    return tables_info


def extract_query_info(access) -> list[dict]:
    """
    Extract information about all queries in the database.
    
    Args:
        access: Active Access Application COM object
        
    Returns:
        List of dictionaries containing query information
    """
    queries_info = []
    
    try:
        current_db = access.CurrentDb()
        query_defs = current_db.QueryDefs
        
        print(f"\nExtracting information from {query_defs.Count} query definitions...")
        
        for query_def in query_defs:
            try:
                query_name = query_def.Name
                
                # Skip system queries (start with ~)
                if query_name.startswith('~'):
                    continue
                
                print(f"  - Analyzing query: {query_name}")
                
                # Get query SQL
                sql = query_def.SQL if hasattr(query_def, 'SQL') else None
                
                # Get query type
                query_type = get_query_type_name(query_def.Type) if hasattr(query_def, 'Type') else 'Unknown'
                
                # Get field information
                fields = []
                try:
                    for field in query_def.Fields:
                        try:
                            fields.append({
                                'name': field.Name,
                                'type': get_field_type_name(field.Type) if hasattr(field, 'Type') else 'Unknown'
                            })
                        except:
                            continue
                except:
                    pass  # Some queries may not have accessible fields
                
                # Get parameters if any
                parameters = []
                try:
                    for param in query_def.Parameters:
                        try:
                            parameters.append({
                                'name': param.Name,
                                'type': get_field_type_name(param.Type) if hasattr(param, 'Type') else 'Unknown'
                            })
                        except:
                            continue
                except:
                    pass
                
                queries_info.append({
                    'name': query_name,
                    'type': query_type,
                    'sql': sql,
                    'fields': fields,
                    'parameters': parameters
                })
                
            except Exception as e:
                print(f"  [ERROR] Error extracting query: {e}")
                continue
        
        print(f"  [OK] Extracted information from {len(queries_info)} queries")
        
    except Exception as e:
        print(f"[ERROR] Error accessing query definitions: {e}")
        import traceback
        traceback.print_exc()
    
    return queries_info


def extract_vba_info(access) -> dict:
    """
    Extract VBA code and module information.
    
    Args:
        access: Active Access Application COM object
        
    Returns:
        Dictionary containing VBA components information
    """
    vba_info = {
        'modules': [],
        'class_modules': [],
        'forms': [],
        'reports': [],
        'total_lines': 0
    }
    
    try:
        vba_project = access.VBE.ActiveVBProject
        
        print(f"\nExtracting VBA code from {vba_project.VBComponents.Count} components...")
        print(f"Project Name: {vba_project.Name}")
        
        for component in vba_project.VBComponents:
            try:
                component_name = component.Name
                component_type = component.Type
                
                # Determine component type
                type_name = {
                    1: "module",           # vbext_ct_StdModule
                    2: "class",            # vbext_ct_ClassModule
                    3: "form",             # vbext_ct_MSForm
                    100: "document"        # vbext_ct_Document (Form/Report)
                }.get(component_type, "unknown")
                
                print(f"  - Extracting: {component_name} ({type_name})")
                
                # Get code from component
                code_module = component.CodeModule
                line_count = code_module.CountOfLines
                
                code_text = ""
                if line_count > 0:
                    code_text = code_module.Lines(1, line_count)
                
                component_info = {
                    'name': component_name,
                    'type': type_name,
                    'type_code': component_type,
                    'line_count': line_count,
                    'code': code_text
                }
                
                # Categorize component
                if component_type == 1:  # Standard module
                    vba_info['modules'].append(component_info)
                elif component_type == 2:  # Class module
                    vba_info['class_modules'].append(component_info)
                elif component_type in [3, 100]:  # Form or Report
                    if "Form_" in component_name or component_type == 3:
                        vba_info['forms'].append(component_info)
                    else:
                        vba_info['reports'].append(component_info)
                
                vba_info['total_lines'] += line_count
                
            except Exception as e:
                print(f"  [ERROR] Error extracting component: {e}")
                continue
        
        print(f"  [OK] Extracted {len(vba_info['modules']) + len(vba_info['class_modules']) + len(vba_info['forms']) + len(vba_info['reports'])} VBA components")
        
    except Exception as e:
        print(f"[ERROR] Error accessing VBA project: {e}")
        import traceback
        traceback.print_exc()
    
    return vba_info


def get_field_type_name(field_type: int) -> str:
    """Convert Access field type code to readable name."""
    type_names = {
        1: "Boolean",
        2: "Byte",
        3: "Integer",
        4: "Long",
        5: "Currency",
        6: "Single",
        7: "Double",
        8: "Date/Time",
        9: "Binary",
        10: "Text",
        11: "OLE Object",
        12: "Memo",
        15: "GUID",
        16: "Big Integer",
        101: "Long Text",
        102: "Calculated",
        103: "Attachment"
    }
    return type_names.get(field_type, f"Unknown ({field_type})")


def get_query_type_name(query_type: int) -> str:
    """Convert Access query type code to readable name."""
    type_names = {
        0: "Select",
        16: "Crosstab",
        32: "Delete",
        48: "Update",
        64: "Append",
        80: "Make-Table",
        96: "DDL",
        112: "SQLPassThrough",
        128: "Union",
        240: "Data Definition"
    }
    return type_names.get(query_type, f"Unknown ({query_type})")


def save_assessment_report(assessment_data: dict, output_dir: str) -> None:
    """
    Save comprehensive assessment report in multiple formats.
    
    Args:
        assessment_data: Dictionary containing all extracted information
        output_dir: Directory where reports will be saved
    """
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Use database name for report files
    db_name = Path(assessment_data['database_name']).stem
    
    # Save JSON report (complete data)
    json_file = output_path / f"{db_name}_assessment.json"
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(assessment_data, f, indent=2, ensure_ascii=False)
    print(f"\n  [OK] Saved JSON report to: {json_file}")
    
    # Save readable text report
    txt_file = output_path / f"{db_name}_assessment.txt"
    with open(txt_file, 'w', encoding='utf-8') as f:
        write_text_report(f, assessment_data)
    print(f"  [OK] Saved text report to: {txt_file}")
    
    # Save markdown report
    md_file = output_path / f"{db_name}_assessment.md"
    with open(md_file, 'w', encoding='utf-8') as f:
        write_markdown_report(f, assessment_data)
    print(f"  [OK] Saved markdown report to: {md_file}")
    
    # Save individual VBA files
    vba_dir = output_path / "vba_code"
    vba_dir.mkdir(exist_ok=True)
    save_vba_files(assessment_data['vba'], vba_dir)
    print(f"  [OK] Saved VBA files to: {vba_dir}")
    
    # Save individual query SQL files
    queries_dir = output_path / "queries"
    queries_dir.mkdir(exist_ok=True)
    save_query_files(assessment_data['queries'], queries_dir)
    print(f"  [OK] Saved query SQL files to: {queries_dir}")


def write_text_report(f, data: dict) -> None:
    """Write assessment data as formatted text report."""
    f.write("="*80 + "\n")
    f.write(f"MS ACCESS DATABASE ASSESSMENT REPORT\n")
    f.write(f"Database: {data['database_path']}\n")
    f.write(f"Assessment Date: {data['assessment_date']}\n")
    f.write("="*80 + "\n\n")
    
    # Summary
    f.write("SUMMARY\n")
    f.write("-"*80 + "\n")
    f.write(f"Tables: {data['summary']['table_count']}\n")
    f.write(f"Queries: {data['summary']['query_count']}\n")
    f.write(f"VBA Modules: {data['summary']['vba_module_count']}\n")
    f.write(f"VBA Lines of Code: {data['summary']['vba_total_lines']}\n")
    f.write("\n\n")
    
    # Tables
    f.write("TABLES\n")
    f.write("="*80 + "\n")
    for table in data['tables']:
        f.write(f"\nTable: {table['name']}\n")
        f.write(f"  Fields: {table['field_count']}\n")
        f.write(f"  Records: {table['record_count'] if table['record_count'] is not None else 'Unknown'}\n")
        
        f.write("\n  Fields:\n")
        for field in table['fields']:
            required = " (Required)" if field.get('required') else ""
            f.write(f"    - {field['name']}: {field['type']}{required}\n")
        
        if table['indexes']:
            f.write("\n  Indexes:\n")
            for index in table['indexes']:
                primary = " (PRIMARY KEY)" if index.get('primary') else ""
                unique = " (UNIQUE)" if index.get('unique') and not index.get('primary') else ""
                fields = ", ".join(index['fields'])
                f.write(f"    - {index['name']}{primary}{unique}: {fields}\n")
        f.write("\n")
    
    # Queries
    f.write("\n\nQUERIES\n")
    f.write("="*80 + "\n")
    for query in data['queries']:
        f.write(f"\nQuery: {query['name']}\n")
        f.write(f"  Type: {query['type']}\n")
        
        if query['parameters']:
            f.write(f"  Parameters:\n")
            for param in query['parameters']:
                f.write(f"    - {param['name']}: {param['type']}\n")
        
        if query['fields']:
            f.write(f"  Fields:\n")
            for field in query['fields']:
                f.write(f"    - {field['name']}: {field['type']}\n")
        
        if query['sql']:
            f.write(f"\n  SQL:\n")
            for line in query['sql'].split('\n'):
                f.write(f"    {line}\n")
        f.write("\n")
    
    # VBA Summary
    f.write("\n\nVBA CODE SUMMARY\n")
    f.write("="*80 + "\n")
    f.write(f"Total Lines: {data['vba']['total_lines']}\n\n")
    
    f.write(f"Standard Modules ({len(data['vba']['modules'])}):\n")
    for module in data['vba']['modules']:
        f.write(f"  - {module['name']}: {module['line_count']} lines\n")
    
    f.write(f"\nClass Modules ({len(data['vba']['class_modules'])}):\n")
    for module in data['vba']['class_modules']:
        f.write(f"  - {module['name']}: {module['line_count']} lines\n")
    
    f.write(f"\nForms with Code ({len(data['vba']['forms'])}):\n")
    for form in data['vba']['forms']:
        f.write(f"  - {form['name']}: {form['line_count']} lines\n")
    
    f.write(f"\nReports with Code ({len(data['vba']['reports'])}):\n")
    for report in data['vba']['reports']:
        f.write(f"  - {report['name']}: {report['line_count']} lines\n")


def write_markdown_report(f, data: dict) -> None:
    """Write assessment data as markdown report."""
    f.write(f"# MS Access Database Assessment Report\n\n")
    f.write(f"**Database:** `{data['database_path']}`  \n")
    f.write(f"**Assessment Date:** {data['assessment_date']}\n\n")
    
    f.write("---\n\n")
    
    # Summary
    f.write("## Summary\n\n")
    f.write(f"- **Tables:** {data['summary']['table_count']}\n")
    f.write(f"- **Queries:** {data['summary']['query_count']}\n")
    f.write(f"- **VBA Modules:** {data['summary']['vba_module_count']}\n")
    f.write(f"- **VBA Lines of Code:** {data['summary']['vba_total_lines']}\n\n")
    
    # Tables
    f.write("## Tables\n\n")
    for table in data['tables']:
        f.write(f"### {table['name']}\n\n")
        f.write(f"- **Fields:** {table['field_count']}\n")
        f.write(f"- **Records:** {table['record_count'] if table['record_count'] is not None else 'Unknown'}\n\n")
        
        f.write("#### Fields\n\n")
        f.write("| Field Name | Type | Required | Default |\n")
        f.write("|------------|------|----------|----------|\n")
        for field in table['fields']:
            required = "Yes" if field.get('required') else "No"
            default = field.get('default_value', '')
            f.write(f"| {field['name']} | {field['type']} | {required} | {default} |\n")
        f.write("\n")
        
        if table['indexes']:
            f.write("#### Indexes\n\n")
            f.write("| Index Name | Type | Fields |\n")
            f.write("|------------|------|--------|\n")
            for index in table['indexes']:
                idx_type = "PRIMARY KEY" if index.get('primary') else ("UNIQUE" if index.get('unique') else "INDEX")
                fields = ", ".join(index['fields'])
                f.write(f"| {index['name']} | {idx_type} | {fields} |\n")
            f.write("\n")
    
    # Queries
    f.write("## Queries\n\n")
    for query in data['queries']:
        f.write(f"### {query['name']}\n\n")
        f.write(f"**Type:** {query['type']}\n\n")
        
        if query['parameters']:
            f.write("**Parameters:**\n")
            for param in query['parameters']:
                f.write(f"- `{param['name']}`: {param['type']}\n")
            f.write("\n")
        
        if query['fields']:
            f.write("**Output Fields:**\n")
            for field in query['fields']:
                f.write(f"- `{field['name']}`: {field['type']}\n")
            f.write("\n")
        
        if query['sql']:
            f.write("**SQL:**\n\n")
            f.write("```sql\n")
            f.write(query['sql'])
            f.write("\n```\n\n")
    
    # VBA
    f.write("## VBA Code\n\n")
    f.write(f"**Total Lines:** {data['vba']['total_lines']}\n\n")
    
    f.write(f"### Standard Modules ({len(data['vba']['modules'])})\n\n")
    if data['vba']['modules']:
        f.write("| Module Name | Lines |\n")
        f.write("|-------------|-------|\n")
        for module in data['vba']['modules']:
            f.write(f"| {module['name']} | {module['line_count']} |\n")
        f.write("\n")
    
    f.write(f"### Class Modules ({len(data['vba']['class_modules'])})\n\n")
    if data['vba']['class_modules']:
        f.write("| Class Name | Lines |\n")
        f.write("|------------|-------|\n")
        for module in data['vba']['class_modules']:
            f.write(f"| {module['name']} | {module['line_count']} |\n")
        f.write("\n")
    
    f.write(f"### Forms with Code ({len(data['vba']['forms'])})\n\n")
    if data['vba']['forms']:
        f.write("| Form Name | Lines |\n")
        f.write("|-----------|-------|\n")
        for form in data['vba']['forms']:
            f.write(f"| {form['name']} | {form['line_count']} |\n")
        f.write("\n")
    
    f.write(f"### Reports with Code ({len(data['vba']['reports'])})\n\n")
    if data['vba']['reports']:
        f.write("| Report Name | Lines |\n")
        f.write("|-------------|-------|\n")
        for report in data['vba']['reports']:
            f.write(f"| {report['name']} | {report['line_count']} |\n")
        f.write("\n")


def save_vba_files(vba_data: dict, output_dir: Path) -> None:
    """Save individual VBA code files."""
    all_components = (
        vba_data['modules'] +
        vba_data['class_modules'] +
        vba_data['forms'] +
        vba_data['reports']
    )
    
    for component in all_components:
        if component['line_count'] == 0:
            continue
        
        # Determine file extension
        if component['type'] == 'module':
            ext = 'bas'
        elif component['type'] == 'class':
            ext = 'cls'
        elif component['type'] == 'form' or 'Form_' in component['name']:
            ext = 'form.bas'
        else:
            ext = 'report.bas'
        
        # Sanitize filename
        safe_name = component['name'].replace('*', '').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
        
        file_path = output_dir / f"{safe_name}.{ext}"
        
        with open(file_path, 'w', encoding='utf-8', errors='replace') as f:
            f.write(f"' Component: {component['name']}\n")
            f.write(f"' Type: {component['type']}\n")
            f.write(f"' Lines: {component['line_count']}\n")
            f.write("' " + "="*60 + "\n\n")
            f.write(component['code'])


def save_query_files(queries_data: list, output_dir: Path) -> None:
    """Save individual query SQL files."""
    for query in queries_data:
        if not query['sql']:
            continue
        
        # Sanitize filename
        safe_name = query['name'].replace('*', '').replace('/', '_').replace('\\', '_').replace(':', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
        
        file_path = output_dir / f"{safe_name}.sql"
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(f"-- Query: {query['name']}\n")
            f.write(f"-- Type: {query['type']}\n")
            
            if query['parameters']:
                f.write("-- Parameters:\n")
                for param in query['parameters']:
                    f.write(f"--   {param['name']}: {param['type']}\n")
            
            f.write("\n")
            f.write(query['sql'])


def assess_access_database(accdb_path: str, output_dir: str) -> dict | None:
    """
    Perform comprehensive assessment of MS Access database.
    
    Args:
        accdb_path: Path to the .accdb or .mdb file
        output_dir: Directory where assessment reports will be saved
        
    Returns:
        Dictionary containing all assessment data, or None if assessment failed
    """
    print(f"Opening Access database: {accdb_path}")
    
    # Validate database file exists
    accdb_path_obj = Path(accdb_path)
    if not accdb_path_obj.exists():
        print(f"Error: Database file not found: {accdb_path}")
        return None
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    # Initialize Access Application
    access = None
    max_retries = 3
    retry_delay = 2
    
    for attempt in range(max_retries):
        try:
            # Clean up any existing Access instances
            if access:
                try:
                    access.CloseCurrentDatabase()
                    access.Quit()
                except:
                    pass
                access = None
                time.sleep(1)
            
            # Create new Access instance
            access = win32com.client.Dispatch("Access.Application")
            
            # Open the database (convert to absolute path)
            abs_path = os.path.abspath(accdb_path)
            
            # Try to open the database
            access.OpenCurrentDatabase(abs_path)
            break  # Success, exit retry loop
            
        except Exception as e:
            error_msg = str(e)
            if "already have the database open" in error_msg.lower() or "already open" in error_msg.lower():
                if attempt < max_retries - 1:
                    print(f"Database appears to be open. Waiting {retry_delay} seconds and retrying... (Attempt {attempt + 1}/{max_retries})")
                    time.sleep(retry_delay)
                    continue
                else:
                    print("\nError: Database is locked by another process.")
                    print("Please close MS Access and any other applications that might have the database open,")
                    print("then try again.")
                    raise
            else:
                # Different error, raise immediately
                raise
    
    try:
        print("\n" + "="*80)
        print("STARTING COMPREHENSIVE DATABASE ASSESSMENT")
        print("="*80)
        
        # Extract all information
        tables = extract_table_info(access)
        queries = extract_query_info(access)
        vba = extract_vba_info(access)
        
        # Build assessment data structure
        assessment_data = {
            'database_path': str(accdb_path_obj.absolute()),
            'database_name': accdb_path_obj.name,
            'assessment_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'summary': {
                'table_count': len(tables),
                'query_count': len(queries),
                'vba_module_count': len(vba['modules']) + len(vba['class_modules']) + len(vba['forms']) + len(vba['reports']),
                'vba_total_lines': vba['total_lines']
            },
            'tables': tables,
            'queries': queries,
            'vba': vba
        }
        
        # Save reports
        print("\n" + "="*80)
        print("SAVING ASSESSMENT REPORTS")
        print("="*80)
        save_assessment_report(assessment_data, output_dir)
        
        # Print summary
        print("\n" + "="*80)
        print("ASSESSMENT SUMMARY")
        print("="*80)
        print(f"Database: {accdb_path_obj.name}")
        print(f"Tables: {assessment_data['summary']['table_count']}")
        print(f"Queries: {assessment_data['summary']['query_count']}")
        print(f"VBA Modules: {assessment_data['summary']['vba_module_count']}")
        print(f"VBA Lines of Code: {assessment_data['summary']['vba_total_lines']}")
        print(f"\nAll reports saved to: {output_dir}")
        
        return assessment_data
        
    except Exception as e:
        print(f"Error during assessment: {e}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        # Close Access
        if access:
            try:
                access.CloseCurrentDatabase()
                access.Quit()
                print("\nAccess closed.")
            except:
                pass


if __name__ == "__main__":
    # Configuration
    script_dir = Path(__file__).parent
    
    # Check for command line arguments
    if len(sys.argv) >= 3:
        accdb_path = sys.argv[1]
        output_dir = sys.argv[2]
    else:
        # Default to TB CMS database
        accdb_path = str(script_dir / "msaccess" / "TB CMS.SQL.accdb")
        output_dir = str(script_dir / "database_assessment")
    
    if not os.path.exists(accdb_path):
        print(f"Error: Access database not found at {accdb_path}")
        print(f"\nUsage: python assess_access_db.py <path_to_accdb> <output_directory>")
        sys.exit(1)
    
    print("="*80)
    print("MS ACCESS DATABASE COMPREHENSIVE ASSESSMENT TOOL")
    print("="*80)
    print()
    
    result = assess_access_database(accdb_path, output_dir)
    
    if result:
        print("\n" + "="*80)
        print("[SUCCESS] Assessment completed successfully!")
        print("="*80)
    else:
        print("\n" + "="*80)
        print("[FAILED] Assessment failed. Please check errors above.")
        print("="*80)
        sys.exit(1)
