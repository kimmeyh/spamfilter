# OutlookMailSpamFilter

Automated Python-based email spam and phishing filter for Microsoft Outlook that processes multiple bulk mail folders using configurable YAML rules.

## Overview

This tool provides intelligent filtering and removal of SPAM and phishing emails from Outlook accounts using pattern-based rules and safe sender management. The system processes emails from configurable folder lists and applies comprehensive filtering criteria including header analysis, body content scanning, subject pattern matching, and sender verification.

## Recent Updates (October 2025)

- ✅ Regex mode is now the default using YAML regex files:
	- Rules: `rulesregex.yaml`
	- Safe senders: `rules_safe_sendersregex.yaml`
- ✅ Legacy files still supported via flag:
	- Rules: `rules.yaml`
	- Safe senders: `rules_safe_senders.yaml`
- ✅ CLI flags added for mode control and one-shot conversions
- ✅ Exporters enforce consistency (lowercase, trimmed, de-duped, sorted) and create timestamped backups in `archive/`
- ✅ Memory bank updated with processing flow, schemas, and regex conventions
- ✅ Interactive prompt gains new options: 'sd' (add sender-domain regex to safe_senders) and '?' (help)

## Recent Updates (July 2025)

- ✅ **Multi-Folder Processing**: Updated to process multiple folders instead of single folder
- ✅ **Configurable Folder List**: EMAIL_BULK_FOLDER_NAMES now supports ["Bulk Mail", "bulk"]
- ✅ **Recursive Folder Search**: Added capability to find folders at any nesting level
- ✅ **Enhanced Logging**: Folder-specific logging and processing information

## Key Features

- **Multi-Folder Processing**: Process emails from configurable list of folders
- **YAML-Based Configuration**: Easy-to-maintain rule files (regex default: `rulesregex.yaml`, `rules_safe_sendersregex.yaml`; legacy: `rules.yaml`, `rules_safe_senders.yaml`)
- **Regex-Default Mode**: Regex YAMLs are used by default; legacy mode available via CLI
- **Multi-Criteria Filtering**: Header, body, subject, and sender-based filtering
- **Phishing Detection**: Suspicious URL and domain analysis
- **Safe Sender Management**: Whitelist trusted senders and domains
- **Comprehensive Logging**: Detailed audit trails of all processing activities
- **Interactive Rule Updates**: Prompts for adding rules based on unmatched emails
	- Options during -u prompts: d/e/s/sd/?
		- d: add domain regex to SpamAutoDeleteHeader (block)
		- e: add full sender email to SpamAutoDeleteHeader (block)
		- s: add literal to safe_senders (allow)
		- sd: add sender-domain regex to safe_senders (allow any subdomain)
		- ?: show brief help
- **Backup System**: Automatic timestamped backups of rule changes
- **Second Pass Reprocessing**: Re-checks remaining emails after interactive updates for additional cleanup

## How to Run

```powershell
# Activate Python virtual environment (PowerShell)
./.venv/Scripts/Activate.ps1
```
```bash
# Activate Python virtual environment (Bash)
source .venv/bin/activate
```
# Run the main application (regex mode is default)
python .\withOutlookRulesYAML.py

# Optional: enable interactive update prompts during the run
python .\withOutlookRulesYAML.py -u

# Force legacy YAML files instead of regex
python .\withOutlookRulesYAML.py --use-legacy-files

# Explicitly use regex files (default behavior)
python .\withOutlookRulesYAML.py --use-regex-files

# One-shot conversions to create/update regex YAMLs from legacy files
python .\withOutlookRulesYAML.py --convert-rules-to-regex
python .\withOutlookRulesYAML.py --convert-safe-senders-to-regex
```

## Configuration

The application targets the **kimmeyharold@aol.com** account and processes emails from the following folders:
- "Bulk Mail"
- "bulk"

Configuration can be modified in the script constants:
- `EMAIL_BULK_FOLDER_NAMES = ["Bulk Mail", "bulk"]`
- `EMAIL_ADDRESS = "kimmeyharold@aol.com"`

## File Structure

- **withOutlookRulesYAML.py** - Main application script
- **rulesregex.yaml** - Regex-mode spam filtering rules (default)
- **rules_safe_sendersregex.yaml** - Regex-mode trusted sender whitelist (default)
- **rules.yaml** - Legacy spam filtering rules
- **rules_safe_senders.yaml** - Legacy trusted sender whitelist
- **requirements.txt** - Python dependencies
- **pytest/** - All test files and test configuration
- **Archive/** - Historical backups and development files
- **memory-bank/** - Configuration for GitHub Copilot memory enhancement

## CLI Flags

- `-u`, `--update_rules`: enable interactive prompts to add header regexes or safe senders during processing
- `--use-regex-files`: use regex YAML files (default behavior)
- `--use-legacy-files`: force legacy YAML files for a run
- `--convert-rules-to-regex`: generate/update `rulesregex.yaml` from `rules.yaml`
- `--convert-safe-senders-to-regex`: generate/update `rules_safe_sendersregex.yaml` from `rules_safe_senders.yaml`

## Testing

All tests are located in the `pytest/` directory. Run tests using:

```bash
# Run all tests
python -m pytest pytest/ -v

# Run specific test file
python -m pytest pytest/test_file_content.py -v
```

Test files include:
- `test_withOutlook_rulesYAML_compare_inport_to_export.py` - YAML import/export validation
- `test_withOutlook_rulesYAML_compare_safe_senders_mport_to_export.py` - Safe senders validation
- `test_file_content.py` - File content validation
- `test_folder_list_changes.py` - Multi-folder configuration tests
- `test_import_compatibility.py` - Import compatibility validation
- `test_second_pass_implementation.py` - Second-pass processing tests

## Dependencies

- win32com.client (Outlook COM interface)
- yaml (YAML file processing)
- logging (Application logging)
- Standard Python libraries (re, datetime, os, etc.)

## Backups and Exporter Invariants

- All list fields (rules conditions/exceptions and safe_senders) are normalized on export:
	- lowercased, trimmed, de-duplicated, and sorted
	- regex YAMLs are written using single quotes to reduce escape noise
- Before overwriting active YAML files, a timestamped backup is created in `archive/`

## Schemas and Conventions

For details, see memory-bank docs:
- `memory-bank/processing-flow.md` — high-level processing, interactive updates, second pass
- `memory-bank/yaml-schemas.md` — effective YAML schemas for rules and safe senders
- `memory-bank/regex-conventions.md` — quoting, glob-to-regex, and domain anchor patterns
	- Includes sender-domain safe-senders regex: `^[^@\s]+@(?:[a-z0-9-]+\.)*<domain>$`
- `memory-bank/quality-invariants.md` — exporter and processing invariants
- `memory-bank/cli-usage.md` — CLI usage reference

## Future Enhancements

- Reprocess emails in multiple folders for additional cleanup passes
- Move backup files to dedicated backup directory
- Regex pattern support for all rule types
- Cross-platform email client support
- Machine learning-based spam detection
