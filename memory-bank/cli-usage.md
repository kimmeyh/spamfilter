# CLI Usage Reference (Updated 11/10/2025)

## Current Mode
**REGEX-only mode** with consolidated YAML filenames (legacy wildcard mode deprecated 10/14/2025)

## Developer Setup
See `memory-bank/dev-environment.md` to create/activate a Python venv before running commands below.

## Running the Application

### Standard Processing (No Interactive Updates)
```powershell
cd D:\Data\Harold\github\OutlookMailSpamFilter && ./.venv/Scripts/Activate.ps1 && python withOutlookRulesYAML.py
```

### Interactive Mode (With Rule Update Prompts)
```powershell
cd D:\Data\Harold\github\OutlookMailSpamFilter && ./.venv/Scripts/Activate.ps1 && python withOutlookRulesYAML.py -u
```

## CLI Flags

### Active Flags
- `-u`, `--update_rules` - Enable interactive prompts to add header regexes or safe senders during processing

### Deprecated Flags (Removed)
- ~~`--use-regex-files`~~ - DEPRECATED 11/10/2025: Regex is now the only mode with consolidated filenames
- ~~`--use-legacy-files`~~ - DEPRECATED 10/14/2025: Legacy wildcard mode removed
- ~~`--convert-rules-to-regex`~~ - DEPRECATED 11/10/2025: Conversion utilities no longer needed
- ~~`--convert-safe-senders-to-regex`~~ - DEPRECATED 11/10/2025: Conversion utilities no longer needed

## Interactive Update Options (-u flag)

When running with `-u` flag, the system prompts for unmatched emails with the following options:

| Option | Action | Description |
|--------|--------|-------------|
| **d** | Block Domain | Add sender domain regex to SpamAutoDeleteHeader (blocks domain and all subdomains) |
| **e** | Block Email | Add full sender email address to SpamAutoDeleteHeader (blocks specific sender) |
| **s** | Allow Literal | Add literal address/domain to safe_senders (never block this specific sender) |
| **sd** | Allow Domain | Add sender-domain regex to safe_senders (never block domain and all subdomains) |
| **?** | Help | Show brief help message with option descriptions |
| **Enter** | Skip | Skip without adding any rules |

### Smart Filtering (10/18/2025)
During interactive sessions, emails matching newly added rules or safe senders are automatically skipped to prevent duplicate prompts for the same domain.

## Entry Points
- `main()` - Primary application entry point
- `OutlookSecurityAgent.set_active_mode()` - Initializes regex mode with consolidated filenames

## File Structure (Consolidated as of 11/10/2025)
- **rules.yaml** - Main spam filtering rules (contains regex patterns)
- **rules_safe_senders.yaml** - Trusted sender whitelist (contains regex patterns)
- **archive/** - Timestamped backups of YAML files created before updates

## Legacy Files (Deprecated)
- ~~rulesregex.yaml~~ → Consolidated to rules.yaml (11/10/2025)
- ~~rules_safe_sendersregex.yaml~~ → Consolidated to rules_safe_senders.yaml (11/10/2025)
