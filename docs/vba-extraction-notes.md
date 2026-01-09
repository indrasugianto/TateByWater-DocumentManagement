# VBA Extraction Domain Knowledge

## MS Access VBA Project Structure

MS Access databases contain VBA projects with multiple component types:

1. **Standard Modules** - Standalone code modules
2. **Class Modules** - Object-oriented class definitions
3. **Forms** - UI forms with code-behind
4. **Reports** - Report definitions with code-behind

## COM Automation Access Pattern

Accessing VBA through COM:

```
Access.Application
  └── VBE (Visual Basic Environment)
      └── ActiveVBProject
          └── VBComponents (collection)
              └── VBComponent
                  ├── Name (string)
                  ├── Type (integer)
                  └── CodeModule
                      ├── CountOfLines (integer)
                      └── Lines(start, count) (string)
```

## Component Type Constants

VBA component types (from `vbext_ComponentType` enum):

- `1` = `vbext_ct_StdModule` - Standard module
- `2` = `vbext_ct_ClassModule` - Class module
- `3` = `vbext_ct_MSForm` - Form
- `100` = `vbext_ct_Document` - Document (Form or Report)

Note: Type 100 can be either a Form or Report. Check component name for `"Form_"` prefix to distinguish.

## Common VBA Patterns

### Option Statements
VBA modules often start with:
```vba
Option Compare Database
Option Explicit
```

### Error Handling
VBA uses `On Error GoTo` pattern:
```vba
On Error GoTo ErrHnd
    ' code here
Exit Sub
ErrHnd:
    ErrMsg "ProcedureName"
End Sub
```

### Public Variables
VBA modules may declare public variables:
```vba
Public dtST1 As Date
Public lngBill_ID As Long
```

## Extraction Challenges

### Empty Components
Some components may have no code (`CountOfLines == 0`). Handle gracefully.

### Encoding Issues
VBA code may contain:
- Legacy character encodings
- Special characters
- Non-ASCII characters

Solution: Use UTF-8 with `errors='replace'`.

### Invalid Filenames
Component names may contain filesystem-invalid characters:
- `<>:"/\|?*`

Solution: Sanitize before creating files.

### COM Object Lifecycle
COM objects must be properly initialized and cleaned up:
1. Create: `Dispatch("Access.Application")`
2. Open database
3. Access VBA project
4. Extract code
5. Close database
6. Quit application

Always use `finally` block for cleanup.

## Performance Notes

- COM operations are relatively slow
- Large databases (100+ components) may take time
- Consider progress indicators for user feedback
- Each component extraction is independent - can parallelize if needed

## Testing Considerations

When testing extraction:

1. Test with various component types
2. Test with empty components
3. Test with components containing special characters
4. Test with large databases
5. Verify file count matches component count
6. Verify file extensions are correct
7. Verify code content is preserved

## Known Limitations

- Requires MS Access to be installed
- Windows-only (COM automation)
- Cannot extract form/report design (only code)
- Cannot extract references/dependencies automatically
- No incremental extraction (always extracts all components)
