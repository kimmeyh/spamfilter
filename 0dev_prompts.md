Next:
Update to consider all Header, Body, Subject, From, lists strings to be regex patterns
create an optional YAML config files for all the major global variables.  List:
EMAIL_BULK_FOLDER_NAMES # list of folders - example ["Bulk Mail", "bulk"] 
EMAIL_INBOX_FOLDER_NAME = "Inbox"
OUTLOOK_SECURITY_LOG_PATH = f"D:/data/harold/OutlookRulesProcessing/"
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



@workspace 



Completed:
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
