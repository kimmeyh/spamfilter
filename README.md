# OutlookMailSpamFilter

Automated Python-based email spam and phishing filter for Microsoft Outlook that processes multiple bulk mail folders using configurable YAML rules.

## Overview

This tool provides intelligent filtering and removal of SPAM and phishing emails from Outlook accounts using pattern-based rules and safe sender management. The system processes emails from configurable folder lists and applies comprehensive filtering criteria including header analysis, body content scanning, subject pattern matching, and sender verification.

## Recent Updates (July 2025)

- ✅ **Multi-Folder Processing**: Updated to process multiple folders instead of single folder
- ✅ **Configurable Folder List**: EMAIL_BULK_FOLDER_NAMES now supports ["Bulk Mail", "bulk"]
- ✅ **Recursive Folder Search**: Added capability to find folders at any nesting level
- ✅ **Enhanced Logging**: Folder-specific logging and processing information

## Key Features

- **Multi-Folder Processing**: Process emails from configurable list of folders
- **YAML-Based Configuration**: Easy-to-maintain rule files (rules.yaml, rules_safe_senders.yaml)
- **Multi-Criteria Filtering**: Header, body, subject, and sender-based filtering
- **Phishing Detection**: Suspicious URL and domain analysis
- **Safe Sender Management**: Whitelist trusted senders and domains
- **Comprehensive Logging**: Detailed audit trails of all processing activities
- **Interactive Rule Updates**: Prompts for adding rules based on unmatched emails
- **Backup System**: Automatic timestamped backups of rule changes

## How to Run

```bash
# Activate Python virtual environment
.venv\scripts\activate

# Run the main application
python .\withOutlookRulesYAML.py
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
- **rules.yaml** - Primary spam filtering rules
- **rules_safe_senders.yaml** - Trusted sender whitelist
- **requirements.txt** - Python dependencies
- **pytest/** - All test files and test configuration
- **Archive/** - Historical backups and development files
- **memory-bank/** - Configuration for GitHub Copilot memory enhancement

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

## Future Enhancements

- Reprocess emails in multiple folders for additional cleanup passes
- Move backup files to dedicated backup directory
- Regex pattern support for all rule types
- Cross-platform email client support
- Machine learning-based spam detection
