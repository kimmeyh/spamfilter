Next:
@workspace use 'memory-bank/*' to understand the workspace 

Any code that should be removed should be commented out and not deleted.
Do not remove any commented out code.
When complete, update the memory-bank/* files and README.md

Assigned to Copilot:
@workspace use 'memory-bank/*' to understand the workspace 
'withOutlookRulesYAM.py' "do NOT use 0dev_prompts.md"
Now we can comment out/deprecate all the functionality for CLI  switch """--use-legacy-files"""
Can you help draft the code for review in the files
Any code that should be removed should be commented out and not deleted.
Do not remove any commented out code.
When complete, update the memory-bank/* files and README.md

Template:
@workspace use 'memory-bank/*' to understand the workspace 
'withOutlookRulesYAM.py' "do NOT use 0dev_prompts.md"

Can you help draft the code for review in the files
Any code that should be removed should be commented out and not deleted.
Do not remove any commented out code.
When complete, update the memory-bank/* files and README.md
----------


create an optional YAML config files for all the major global variables.  List:
EMAIL_BULK_FOLDER_NAMES # list of folders - example ["Bulk Mail", "bulk"] 
EMAIL_INBOX_FOLDER_NAME = "Inbox"
OUTLOOK_SECURITY_LOG_PATH = f"D:/Data/Harold/OutlookRulesProcessing/"
OUTLOOK_SECURITY_LOG = OUTLOOK_SECURITY_LOG_PATH + "OutlookRulesProcessingDEBUG_INFO.log"
OUTLOOK_SIMPLE_LOG = OUTLOOK_SECURITY_LOG_PATH + "OutlookRulesProcessingSimple.log"
OUTLOOK_RULES_PATH = f"D:/data/harold/github/OutlookMailSpamFilter/"
OUTLOOK_RULES_FILE = OUTLOOK_RULES_PATH + "outlook_rules.csv"
OUTLOOK_SAFE_SENDERS_FILE = OUTLOOK_RULES_PATH + "OutlookSafeSenders.csv"
YAML_RULES_PATH = f"D:/data/harold/github/OutlookMailSpamFilter/"
YAML_ARCHIVE_PATH = YAML_RULES_PATH + "archive/"
YAML_RULES_FILE = YAML_RULES_PATH + "rules.yaml"
#YAML_RULES_FILE = YAML_RULES_PATH + "rules_new.yaml" # this was temporary and no longer needed
YAML_RULES_SAFE_SENDERS_FILE    = YAML_RULES_PATH + "rules_safe_senders.yaml"

# not sure if these will be used
YAML_RULES_BODY_FILE            = YAML_RULES_PATH + "rules_body.yaml"
YAML_RULES_HEADER_FILE          = YAML_RULES_PATH + "rules_header.yaml"
YAML_RULES_SUBJECT_FILE         = YAML_RULES_PATH + "rules_subject.yaml"
YAML_RULES_SPAM_FILTER_FILE     = YAML_RULES_PATH + "rules_spam_filter.yaml"
YAML_RULES_SAFE_RECIPIENTS_FILE = YAML_RULES_PATH + "rules_safe_recipients.yaml"
YAML_RULES_BLOCKED_SENDERS_FILE = YAML_RULES_PATH + "rules_blocked_senders.yaml"
YAML_RULES_CONTACTS_FILE        = YAML_RULES_PATH + "rules_contacts.yaml"           # periodically review email account contacts and update
YAML_RULES_EMAIL_TO_FILE        = YAML_RULES_PATH + "rules_email_to.yaml"           # periodically review emails sent and add targeted recipients to secondary "Safe Senders" file (name?)
YAML_INTERNATIONAL_RULES_FILE   = YAML_RULES_PATH + "rules_international.yaml"      # send all but a few "organizations" "*.<>" to Bulk Mail .jp, .cz...
OUTLOOK_RULES_SUBSET            = "SpamAutoDelete"
DAYS_BACK_DEFAULT = 365 # default number of days to go back in the calendar

Often needed:
Can you review all the memory-bank/ files and ensure they are current based on the contents of withOutlookRulesYAML.py and the REGEX files: rulesregex.yaml and rules_safe_senderregex.yaml
Propose updates to memory-bank/*.* files

How to Run from command line (best practice):
cd D:\Data\Harold\github\OutlookMailSpamFilter && ./.venv/Scripts/Activate.ps1 && python withOutlookRulesYAML.py -u
cd D:\Data\Harold\github\OutlookMailSpamFilter && ./.venv/Scripts/Activate.ps1 && python withOutlookRulesYAML.py

------------------------------------------------------------------------------
Completed:
'withOutlookRulesYAM.py' "do NOT use 0dev_prompts.md"
The spam filtering is working as expected, except during the user input. For example, if the user enters "d" to add the domain rule, it should add that rule so that any future occurrences of that domain are filtered before input is requested.  Same for "sd".  Can you help

for the "sd" input value, it looks like it has been adding the top-level domain, first sub-domain and second sub-domain.  Can you 
help adjust so that it only includes the top-level domain, and first sub-domain.
Example input:  something@mail.cursor.com resulted in - '^[^@\s]+@(?:[a-z0-9-]+\.)*mail\.cursor\.com$' but
should result in - '^[^@\s]+@(?:[a-z0-9-]+\.)*cursor\.com$'
Can you help fix.

Is there a way to complete the following Outlook Classic Client menu process in the codebase:
Home > Delete > Junk > Never Block Sender's Domain 

During user input for "Add <> to SpamAutoDeleteHeader rule or safe_senders? (d/e/s):" can you add 
and then implement an additional response "sd" for Add "Senders Domain".
Implementation should be adding a regex similar to example '^[^@\s]+@(?:[a-z0-9-]+\.)*lifeway\.com$'
Where it includes the top-level domain (.com) and sub-level domain (lifeway), with any number of prior sub-domains and any email
name.  Ask if you have any questions.

It does not appear to be using either the rulesregex.yaml or rules_safe_sendersreges.yaml as requested via 
CLI content or it is not using the same logic on the second pass for processing regex patterns
Can you check and identify why this is happening?

Just ran 'withOutlookRulesYAM.py' and it did not match the following from address to regex in rulesregxx.yaml 
Can you help me understand why and fix so that it does

Can you help me guide me on the order for doing the following upgrade to the withOutlookRulesYAM.py:
- convert rules.yaml from existing rules (if you look at the first 20 rows of each rule, the would be sufficient to understand) to equivalent regex strings. Likely at about 500 rule entries at a time
- convert rules_safe_senders.yaml from existing rules (if you look at the first 20 rows of each rule, the would be sufficient to understand) to equivalent regex strings. Likely at about 500 rule entries at a time
- convert the processing in 'withOutlookRulesYAM.py'of the rules.yaml rules so that it now expects regex strings as input and processing the rules as regex strings.
- convert the processing in 'withOutlookRulesYAM.py'of the rules_safe_senders.yaml rules so that it now expects regex strings as input and processing the rules as regex strings.
- for intermediate/small adjustments may want to add the rules as new files (with _regex) at the end of the filenames and an option to use either the current or _regex version of the files/processing in order to incrementally test changes.
- provide pytests to insure all changes are working as expected.
- provide a way to back out changes to the current working version if changes do not work as expected

Can you help me add an input parameter, -update_rules or -u, to toggle if prompt_update_rules is called.  it should default to not call prompt_update_rules

I want you to review file rulesregex.yaml for yaml errors and fix them.
from past experience, because the lists of strings under header, body, subject, and from strings are complex regex strings most must be single quoted.  For consistency keep all lines with single quotes.

✓ Reprocess all emails in the EMAIL_BULK_FOLDER_NAMES folder list a second time, in case any of the remaining emails can no be moved or deleted.

✓ Change EMAIL_BULK_FOLDER_NAME from single folder name to list of folders, add "bulk", ONLY change code that HAS to be CHANGED
cam you help me setup the memory-bank mcp server (what files does it rely on, can you create the files, update them based on workspace content...)

Update mail processing to use safe_senders list for all header exceptions

please write an algorithm to export the current json_rules (see rest of program for reference
to what the JSON looks like and example file @rules.yaml).
The rules.yaml needs to be accurate in output and format.
Then update get_yaml_rules, to read in from YAML_RULES_FILE so that it exactly match json_rules prior ot export.
ONLY change code that HAS to be CHANGED to implement the recommendation.
Any code that should be removed should be commented out and not deleted.
Do not remove any commented out code.

Can you update get_safe_senders_rules to read in the safe_senders file and make a separate JSON variable for the

Can you create a protobuf schema for rules_new.yaml
