# New Command-Line Flags - Semantics and Mutual Exclusivity

## Overview

Three new mutually exclusive flags have been added to provide different processing modes for document handling:

## Flag Definitions

### `--fix-errors`
- **Purpose**: Run spelling and grammar corrections automatically
- **Behavior**: 
  - Implies `--check-spelling` but with automatic fixing
  - Shows initial error report
  - Applies automatic corrections (implementation pending)
  - Continues with document formatting after corrections
- **Relationship to existing flags**: Supersedes `--check-spelling` with auto-fix capability

### `--report-only` 
- **Purpose**: Generate comprehensive reports without making any changes
- **Behavior**:
  - Generates spelling and grammar error reports
  - Generates formatting analysis reports
  - **Does NOT** format, fix, or modify the document in any way
  - Exits after report generation
- **Relationship to existing flags**: Mutually exclusive with formatting and fixing operations

### `--just-format`
- **Purpose**: Apply citation and style formatting only
- **Behavior**:
  - Skips all spelling and grammar checks
  - Applies only citation/style formatting rules
  - Faster processing by bypassing text analysis
- **Relationship to existing flags**: Bypasses `--check-spelling` and `--spell-only`

## Mutual Exclusivity

The three new flags are enforced as mutually exclusive using `argparse.add_mutually_exclusive_group()`:

```python
mode_group = parser.add_mutually_exclusive_group()
mode_group.add_argument("--fix-errors", ...)
mode_group.add_argument("--report-only", ...)
mode_group.add_argument("--just-format", ...)
```

**Cannot be used together:**
- `--fix-errors` + `--report-only` 
- `--fix-errors` + `--just-format`
- `--report-only` + `--just-format`

## Relationship to Existing Flags

### Compatible with new flags:
- `-s/--style`: All new flags respect the chosen citation style
- `-o/--output`: Output path specification works with `--fix-errors` and `--just-format`
- `-i/--interactive`: Interactive mode works with all new flags
- `-l/--list-styles`: Independent functionality

### Superseded by new flags:
- `--check-spelling`: Superseded by `--fix-errors` (with auto-fix) and `--report-only` (report mode)
- `--spell-only`: Superseded by `--report-only` (includes spelling + formatting analysis)

## Usage Examples

```bash
# Generate comprehensive report without changes
python app.py document.docx --report-only

# Auto-fix errors and format
python app.py document.docx --fix-errors --style mla

# Format only, skip spell check
python app.py document.docx --just-format --style chicago

# Invalid - will show error
python app.py document.docx --fix-errors --report-only
```

## Implementation Notes

- The mutual exclusivity is enforced at the argument parsing level
- `--report-only` exits early and prevents any document modifications
- `--fix-errors` continues to formatting after applying corrections
- `--just-format` bypasses all text analysis for faster processing
- All flags maintain compatibility with existing style and output options
