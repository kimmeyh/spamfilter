# 02/15/25 Harold Kimmey Add YAML rules exports of outlook rules
# 02/15/25 Harold Kimmey All working as expected
# 01/24/25 Harold Kimmey All working as expected
# 02/10/2025 Harold Kimmey Completeed move to www.github.com/kimmeyh/spamfilter.git repository
# 02/15/2025 Harold Kimmey Start YAML rules exports of outlook rules to YAML files
# 02/17/2025 Harold Kimmey Updated _process_actions to accurately pull the assign_to_category value by searching it as an object
# 03/30/2025 Harold Kimmey All working as expected - commit and PR
# I've modified the security agent to specifically target the "Bulk Mail" folder in the kimmeyharold@aol.com account. Key changes include:

# 1. Account/Folder Targeting:
#    - Added account-specific folder lookup
#    - Recursive folder search capability (in case "Bulk Mail" is nested)
#    - Validation of account and folder existence

# 2. Improved Structure:
#    - Creates "Security Review" folder within the Bulk Mail folder
#    - Only processes emails from the specified folder
#    - Maintains all security checks and rule processing

# 3. Better Error Handling:
#    - Validates account and folder access
#    - Detailed logging of folder navigation
#    - Graceful handling of missing folders

# To use this version:
# 1. Make sure Outlook is running and the AOL account is connected
# 2. Run the script - it will automatically target the Bulk Mail folder
# 3. Suspicious emails will be moved to a "Security Review" subfolder within Bulk Mail

#--------------------------------------------
# 03/28/2025 Harold Kimmey Exported rules to JSON so that they can be maintained in a separate YAML file (the can be transferred between machines and platforms (Windows, Mac, Linux, Android, iOS, etc.))
#   Spam filter rules - done
#   Safe Senders - done
#   Safe recipients - done (was empty)
#   Blocked Senders - very small and were added manually to Outlook Rules - spam body
# 03/29/2025 Harold Kimmey Add functionality to update YAML rules near the end when new domains are listed
#       Simple Y/N to add to the "@<domain>" to the header rules in the YAML file
#       Then make change to use the YAML file as the default (no longer outlook)

#List of future enhancements
#     body, header, subject rules

# Temporary - before writing rules, add safe recipients from CSV
# Not sure if I have a full list of OK'd domains and users.  May need to look through emails sent and stored and add - for
#   when I will go directly against mail.aol.com and need to move messages to Bulk or leave in the inbox

#   Contacts - will need to write this
# Temporary - before writing rules, add safe recipients from CSV
#      can remove once reliant on YAML rules
    #   implement: disable rules for spam and phishing rules before deleting
    #   Warning about suspicious domain names and email addresses
    #   safe sender domains and email addresses (pull from header "From")
    #       use after @ and match domain exactly, but partials, ex. %.microsoft.com...
    #   import contacts and trust email from contacts
    #   implement flag for automatically add people I email to safe senders list
    #   implement "Blocked Senders" or incorporate into header rules
    #   implement international rules (block all but a few "organizations" "*.jp" to Bulk Mail)
# Convert all rules interpretation to be exact strings unless regex content '(i?)" is found, F
#    then assume they are regex
#
# Add "easy to add to Outlook Rules"
#  - Track all the "No conditions or phishing indicators found" as you go by Outlook Rule:
#     - so that you can write a summary at the end, keep a record for each of From:, Subject:, Body:, Header:
#  - Then list summary at the end:
#    - make it easy to copy/paste into Outlook Rules, one rule at a time.
#   Body - then one line per with .<domain>. and /<domain>.
#   From - one line per with @<domain>.
#   Subject/From - one line with From: trimmed to "@<domain>.", Subject: <subject>
# Add ability to do auto-updates to the Outlook Rules for SpamAutoDeleteBody to list of Conditions_obj.Body.Text: add both .<domain>. and /<domain>.
# Add ability to do auto-updates to the Outlook Rule for SpamAutoDeleteHeader to list of Conditions_obj.MessageHeader.Text: from From: trimmed to "@<domain>."
#
# ----------------------------------------------------
# (not in this order, probably later) Convert from using win32com to using o365
#
# Successfully export rules to a CSV file that logically matches the JSON object at the end of the run
# Successfully import rules from the CSV file at the beginning of the run that matches the JSON object from get_outlook_rules
# Start to use the CSV file as the primary source of rules
# Add logic to add items to the rules JSON object via user input at the end of the run
#   - verify that the CSV file is updated with the new rules and can read them back in successfully
# Move all the appropriate rules to a CSV file structure
#   read from CSV and convert back to JSON object
#   add field to rules JSON for outlook "flag" to be applied
#   add rules from outlook for Junk Email "Safe Senders", "Safe Recipients" and "Blocked Snders"
# For Outlook only - may have to async the delete and try 10 times with 1 second delay - get around "can't delete because message has been changed error"
# Change to a phone app that processes emails from cloud provider email accounts:  aol, gmail, yahoo, etc.
#   - in a language that can be used on all platforms:  Android, iOS, Windows, Mac, Linux
#   - use the same JSON rules file for all platforms
#   - use the same JSON rules file for all email accounts
#   - use the same JSON rules and allow uniqueness for different email accounts and account/folder combinations
#   - allow for multiple email accounts to be processed
#   - allow for multiple folders to be processed
#   - Allow options similar to Outlook Junk options: junk level (Safe Lists Only/High/Low/No automatic filtering - see Outlook window
#   - Allow for option to notify of emails with suspicious domain names in email addresses and links in the body
#   - Add a curated, updated list of suspicious domain names, why, level...
#   - Parameterize some of the variables so they can be:
#       - Saved to a file
#       - Read from a file
#       - Updated by a standars process: OUTLOOK_RULES_SUBSET,
#       OUTLOOK_RULES_PATH, OUTLOOK_RULES_FILE, EMAIL_ADDRESS, EMAIL_FOLDER_NAME, OUTLOOK_SECURITY_LOG,
#       OUTLOOK_SIMPLE_LOG, DAYS_BACK_DEFAULT, DEBUG_EMAILS_TO_PROCESS...
# Add support for multiple folders?
# Implement different processing rules for different folders?
# Add email volume reporting?
# Create a summary report of processed emails?

#Imports for python base packages
import re
from datetime import datetime, timedelta
import logging
import sys
import json
import os
import csv
import yaml
import winreg

#Imports for packages that need to be installed
import win32com.client
import IPython

# Settings:
DEBUG = False #True
INFO = False if DEBUG else True #If not debugging, then INFO level logging
DEBUG_EMAILS_TO_PROCESS = 10000 #100 for testing

CRLF = "\n"
EMAIL_ADDRESS = "kimmeyharold@aol.com"
EMAIL_FOLDER_NAME = "Bulk Mail"
WIN32_CLIENT_DISPATCH = "Outlook.Application"
OUTLOOK_GETNAMESPACE = "MAPI"
OUTLOOK_SECURITY_LOG_PATH = f"D:/data/harold/OutlookRulesProcessing/"
OUTLOOK_SECURITY_LOG = OUTLOOK_SECURITY_LOG_PATH + "OutlookRulesProcessingDEBUG_INFO.log"
OUTLOOK_SIMPLE_LOG = OUTLOOK_SECURITY_LOG_PATH + "OutlookRulesProcessingSimple.log"
OUTLOOK_RULES_PATH = f"D:/data/harold/github/OutlookMailSpamFilter/"
OUTLOOK_RULES_FILE = OUTLOOK_RULES_PATH + "outlook_rules.csv"
OUTLOOK_SAFE_SENDERS_FILE = OUTLOOK_RULES_PATH + "OutlookSafeSenders.csv"
YAML_RULES_PATH                 = f"D:/data/harold/OutlookRulesProcessing/"
YAML_RULES_ALL_FILE             = YAML_RULES_PATH + "rules_all.yaml"
# not sure if these will be used
YAML_RULES_BODY_FILE            = YAML_RULES_PATH + "rules_body.yaml"
YAML_RULES_HEADER_FILE          = YAML_RULES_PATH + "rules_header.yaml"
YAML_RULES_SUBJECT_FILE         = YAML_RULES_PATH + "rules_subject.yaml"
YAML_RULES_SPAM_FILTER_FILE     = YAML_RULES_PATH + "rules_spam_filter.yaml"
YAML_RULES_SAFE_SENDERS_FILE    = YAML_RULES_PATH + "rules_safe_senders.yaml"
YAML_RULES_SAFE_RECIPIENTS_FILE = YAML_RULES_PATH + "rules_safe_recipients.yaml"
YAML_RULES_BLOCKED_SENDERS_FILE = YAML_RULES_PATH + "rules_blocked_senders.yaml"
YAML_RULES_CONTACTS_FILE        = YAML_RULES_PATH + "rules_contacts.yaml"           # periodically review email account contacts and update
YAML_RULES_EMAIL_TO_FILE        = YAML_RULES_PATH + "rules_email_to.yaml"           # periodically review emails sent and add targeted recipients to secondary "Safe Senders" file (name?)
YAML_INTERNATIONAL_RULES_FILE   = YAML_RULES_PATH + "rules_international.yaml"      # send all but a few "organizations" "*.<>" to Bulk Mail .jp, .cz...

OUTLOOK_RULES_SUBSET = "SpamAutoDelete"
DAYS_BACK_DEFAULT = 365 #default number of days to go back in the calendar


def simple_print(message):
    """Print message to a file or stdout based on OUTLOOK_SIMPLE_LOG"""
    if OUTLOOK_SIMPLE_LOG:
        with open(OUTLOOK_SIMPLE_LOG, 'a') as f:
            f.write(message + '\n')
    else: #write to the console
        print(message)

def export_rules_yaml (rules_json, rules_yaml):  # try without executing
    """Export rules to a YAML file"""
    import yaml

    try:
        with open(rules_yaml, 'w') as yaml_file:
            yaml.dump(rules_json, yaml_file, default_flow_style=False)
        simple_print(f"Rules exported to {rules_yaml}")
    except Exception as e:
        simple_print(f"Error exporting rules to {rules_yaml}: {str(e)}")

def import_rules_yaml(rules_yaml):
    """Import rules from a YAML file"""
    import yaml
    try:
        with open(rules_yaml, 'r') as yaml_file:
            rules_json = yaml.safe_load(yaml_file)
        print(f"Rules imported from {rules_yaml}")
        return rules_json
    except Exception as e:
        print(f"Error importing rules from {rules_yaml}: {str(e)}")
        return None

class OutlookSecurityAgent:
    def __init__(self, email_address=EMAIL_ADDRESS, folder_name=EMAIL_FOLDER_NAME, debug_mode=DEBUG):
        """
        Initialize the Outlook Security Agent with specific account and folder

        Args:
            email_address: Email address of the account to process
            folder_name: Name of the folder to process
            debug_mode: If True, run in simulation mode with verbose output
        """
        self.debug_mode = debug_mode
        self.outlook = win32com.client.Dispatch(WIN32_CLIENT_DISPATCH)
        self.namespace = self.outlook.GetNamespace(OUTLOOK_GETNAMESPACE)

        # Configure logging
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        logging.basicConfig(
            level=logging.DEBUG if debug_mode else logging.INFO,
            format=log_format,
            handlers=[
                logging.FileHandler(OUTLOOK_SECURITY_LOG),
                # logging.StreamHandler(sys.stdout)  # Also print to console
            ]
        )
        self.log_print(f"Starting new run\n=============================================================")
        self.log_print(f"Initializing agent for {email_address}, folder: {folder_name}")
        self.log_print(f"Debug mode: {debug_mode}")

        # Get the specific account's folder
        self.target_folder = self._get_account_folder(email_address, folder_name)
        if not self.target_folder:
            self.log_print(f"Could not find folder '{folder_name}' in account '{email_address}'")
            raise ValueError(f"Could not find folder '{folder_name}' in account '{email_address}'")

        self.rules = []
        self.rule_to_category = {
            "SpamAutoDeleteBody":           "SpamBody",
            "SpamAutoDeleteBody-imgur.com": "SpamImgur",
            "SpamAutoDeleteFrom":           "SpamHeader",
            "SpamAutoDeleteHeader":         "SpamHeader",
            "SpamAutoDeleteSubject":        "SpamSubject"
        }

    def log_print(self, message, level="INFO"):
        try:
            sanitized_message = self._sanitize_string(message)
            logging.debug(sanitized_message) if level == "DEBUG" else None
            logging.info(sanitized_message) if level == "INFO" else None
        except UnicodeEncodeError:
            logging.debug(sanitized_message.encode('utf-8', 'replace').decode('utf-8')) if level == "DEBUG" else None
            logging.info(sanitized_message.encode('utf-8', 'replace').decode('utf-8'))  if level == "INFO" else None

    def _sanitize_string(self, s):
        """Sanitize string to replace non-ASCII characters"""
        try:
            return re.sub(r'[^\x00-\x7F]+', '', s)
        except UnicodeEncodeError:
            return re.sub(r'[^\x00-\x7F]+', '', s.encode('utf-8', 'replace').decode('utf-8'))

    def _get_account_folder(self, email_address, folder_name):
        """Get a specific folder from a specific email account"""
        self.log_print(f"Searching for folder: {folder_name} in account: {email_address}", "DEBUG")

        try:
            # Loop through accounts to find the matching one
            for account in self.outlook.Session.Accounts:
                self.log_print(f"Checking account: {account.SmtpAddress}", "DEBUG")

                if account.SmtpAddress.lower() == email_address.lower():
                    self.log_print(f"Found matching account: {account.SmtpAddress}")

                    # Get the root folder for this account
                    root_folder = self.namespace.Folders(account.DeliveryStore.DisplayName)
                    self.log_print(f"Accessed root folder: {root_folder.Name}", "DEBUG")

                    # Search for the target folder
                    try:
                        # Try direct access first
                        target_folder = root_folder.Folders[folder_name]
                        self.log_print(f"Found target folder directly: {folder_name}")
                        return target_folder
                    except Exception:
                        self.log_print(f"Folder not found directly, searching recursively...")
                        return self._find_folder_recursive(root_folder, folder_name)

            self.log_print(f"Account not found: {email_address}")
            return None

        except Exception as e:
            self.log_print(f"Error finding account folder: {str(e)}")
            return None

    def _escape_pattern(self, value):
        """Escape special characters in values for CSV storage"""
        if not isinstance(value, str):
            return value, False
        needs_special = False
        if any(char in value for char in [',', '"', "'", '\\', '\n', '\r', ';']):
            needs_special = True
            value = value.replace('\\', '\\\\').replace('"', '\\"')
        return value, needs_special

    def _unescape_pattern(self, value):
        """Unescape value from CSV storage"""
        if not isinstance(value, str):
            return value
        return value.replace('\\"', '"').replace('\\\\', '\\')

    def compare_rules(self, rules1, rules2):
        """Compare two sets of rules and return the differences."""
        # Convert single rules to lists if needed
        rules1_list = [rules1] if isinstance(rules1, dict) else rules1
        rules2_list = [rules2] if isinstance(rules2, dict) else rules2

        # Create dictionaries keyed by rule name for easy comparison
        rules1_dict = {}
        rules2_dict = {}

        # Safely create dictionaries with error handling
        for rule in rules1_list:
            if isinstance(rule, dict) and 'name' in rule:
                rules1_dict[rule['name']] = rule
            else:
                self.log_print(f"Warning: Invalid rule format in first set: {rule}")

        for rule in rules2_list:
            if isinstance(rule, dict) and 'name' in rule:
                rules2_dict[rule['name']] = rule
            else:
                self.log_print(f"Warning: Invalid rule format in second set: {rule}")

        # Find rules unique to each set
        rules_only_in_1 = set(rules1_dict.keys()) - set(rules2_dict.keys())
        rules_only_in_2 = set(rules2_dict.keys()) - set(rules1_dict.keys())

        # Find modified rules (present in both but different)
        modified_rules = {}
        common_rules = set(rules1_dict.keys()) & set(rules2_dict.keys())
        for rule_name in common_rules:
            if rules1_dict[rule_name] != rules2_dict[rule_name]:
                modified_rules[rule_name] = {
                    'rules1': rules1_dict[rule_name],
                    'rules2': rules2_dict[rule_name]
                }

        return {
            'rules_only_in_1': [rules1_dict[name] for name in rules_only_in_1],
            'rules_only_in_2': [rules2_dict[name] for name in rules_only_in_2],
            'modified_rules': modified_rules
        }

    def output_rules_differences(self, outlook_rules, CSV_rules):
        """Output the differences between CSV_rules and outlook_rules"""
        differences = self.compare_rules(outlook_rules, CSV_rules)

        # Print the differences
        if differences['rules_only_in_1']:
            self.log_print("\nRules only in outlook_rules:")
            for rule in differences['rules_only_in_1']:
                self.log_print(f"- {rule['name']}")

        if differences['rules_only_in_2']:
            self.log_print("\nRules only in CSV_Rules set:")
            for rule in differences['rules_only_in_2']:
                self.log_print(f"- {rule['name']}")

        if differences['modified_rules']:
            self.log_print("\nModified rules:")
            for name, rules in differences['modified_rules'].items():
                self.log_print(f"- {name} has differences")

        return

    # def get_outlook_junk_mail_options(self):
    #     """
    #     Retrieve the Outlook Junk Email Options settings (as shown in Outlook Classic > Home > Junk Email Options > Options)
    #     and convert them to a dictionary for further processing or export.
    #     """
    #     timestamp = datetime.now().isoformat()
    #     try:
    #         # Access the JunkEmailOptions directly from the DefaultStore
    #         options = self.outlook.Session.DefaultStore.JunkEmailOptions
    #         # Build a dictionary with key properties.
    #         # (Property names may vary between Outlook versions.)
    #         options_dict = {
    #             'last_modified' : timestamp,
    #             'name'          : "JunkEmailOptions",
    #             'filter_level'  : getattr(options, 'FilterLevel', None),  # e.g., 0=Off, 1=Low, 2=High (depending on your Outlook version)
    #             'enabled'       : getattr(options, 'Enabled', True),  # Some versions may provide an Enabled property
    #             # The safe and blocked lists are typically collections. Convert them to lists if available.
    #             'safe_senders'  : list(options.SafeSenders) if getattr(options, 'SafeSenders', None) else [],
    #             'blocked_senders': list(options.BlockedSenders) if getattr(options, 'BlockedSenders', None) else [],
    #             # Optional: if your Outlook exposes domains lists:
    #             'safe_domains'  : list(options.SafeDomains) if hasattr(options, 'SafeDomains') and options.SafeDomains else [],
    #             'blocked_domains': list(options.BlockedDomains) if hasattr(options, 'BlockedDomains') and options.BlockedDomains else [],
    #         }
    #     except Exception as e:
    #         self.log_print(f"Error processing Junk Email Options: {str(e)}")
    #         return {}

    #     if DEBUG:
    #         self.log_print(f"Junk Email Options retrieved: {options_dict}")
    #     return options_dict # will need to be converted and appended to the json rules object


    # NOTE: tried to get the outlook junk email options and lists, but could not get it to work

    def get_outlook_rules(self):
        """
        Convert Outlook rules to JSON format with comprehensive error checking.
        Returns a list of rule dictionaries with all available properties.
        """
        self.log_print(f"{CRLF}Converting Outlook rules to JSON format...")
        rules_json = []
        rules_dict = {}
        timestamp = datetime.now().isoformat()

        try:
            # NOTE: GetRules() is not returning several of the actions:
            #   - Mark as Read
            #   - Clear the Message Flag
            #   - Stop Processing More Rules
            #   Also, the "AssignToCategory" is not returning the category name

            # Get all rules that start with the subset name
            outlook_rules = [rule for rule in self.outlook.Session.DefaultStore.GetRules() if rule.Name.startswith(OUTLOOK_RULES_SUBSET)]

            self.log_print(f"Processing {len(outlook_rules)} rules...")

            for rule in outlook_rules:
                try:
                    self.log_print(f"\n\nAnalyzing rule: {rule.Name}")
                    rule_dict = {
                        'last_modified': timestamp,
                        "name": rule.Name if hasattr(rule, "Name") else "Unnamed Rule",
                        "enabled": bool(rule.Enabled) if hasattr(rule, "Enabled") else False,
                        "isLocal": bool(rule.IsLocalRule) if hasattr(rule, "IsLocalRule") else False,
                        "executionOrder": rule.ExecutionOrder if hasattr(rule, "ExecutionOrder") else 0,
                        "conditions": {},
                        "actions": {},
                        "exceptions": {},
                    }

                    # Process Conditions
                    if hasattr(rule, "Conditions") and rule.Conditions:
                        conditions = rule.Conditions
                        rule_dict["conditions"] = self._process_conditions(conditions, False)

                    # Process Actions
                    if hasattr(rule, "Actions") and rule.Actions:
                        actions = rule.Actions
                        rule_dict["actions"] = self._process_actions(actions)

                    # Process Exceptions
                    if hasattr(rule, "Exceptions") and rule.Exceptions:
                        exceptions = rule.Exceptions
                        rule_dict["exceptions"] = self._process_conditions(exceptions, True)  # Exceptions use same format as conditions

                    rules_json.append(rule_dict)
                    self.log_print(f"Successfully processed rule: {rule_dict['name']}", "DEBUG")

                except Exception as e:
                    self.log_print(f"Error processing rule {getattr(rule, 'Name', 'Unknown')}: {str(e)}")
                    # Add error information to the rule
                    rules_json.append({
                        "name": getattr(rule, "Name", "Unknown Rule"),
                        "error": str(e),
                        "processed": False
                    })

            # call here once debugged and working
            # temp = self.get_outlook_junk_email_options()

            # print (json.dumps(rules_json, indent=2, default=str)) #can be used for extra debugging information
            return json.loads(json.dumps(rules_json, indent=2, default=str))

        except Exception as e:
            self.log_print(f"Error accessing Outlook rules: {str(e)}")
            return json.dumps({"error": str(e)})


    def import_rules(self, rules_file=OUTLOOK_RULES_FILE):
        """Import rules from CSV file and return as JSON object (not string)"""
        self.log_print("Converting CSV rules to JSON format...")
        try:
            if not os.path.exists(rules_file):
                self.log_print("Rules CSV file not found")
                return None

            rules_dict = {}
            timestamp = datetime.now().isoformat()

            with open(rules_file, 'r', encoding='utf-8', newline='') as f:
                reader = csv.DictReader(f)

                for row in reader:
                    rule_name = row['rule_name']

                    # Initialize rule if we haven't seen it before
                    if rule_name not in rules_dict:
                        rules_dict[rule_name] = {
                            'name': rule_name,
                            'enabled': row['enabled'].lower() == 'true',
                            'isLocal': True,
                            'executionOrder': int(row['execution_order']),
                            'conditions': {
                                'from': [],
                                'subject': [],
                                'body': [],
                                'header': []
                            },
                            'actions': {
                                'move_to_folder': {'folder_path': None, 'folder_name': None},
                                'copy_to_folder': {'folder_path': None, 'folder_name': None},
                                'assign_to_category': {'category_name': None},
                                'delete': row['action_delete'].lower() == 'true',
                                'forward': [],
                                'redirect': [],
                                'play_sound': {'sound_file': None},
                                'mark_as_task': {'task_due_date': None}
                            },
                            'exceptions': {
                                'from': [],
                                'subject': [],
                                'body': [],
                                'header': []
                            },
                            'last_modified': timestamp  # Add last_modified with default timestamp
                        }

                    # Get the value and unescape if needed
                    value = row['condition_value']
                    if row['needs_special_handling'].lower() == 'true':
                        value = self._unescape_pattern(value)

                    # Add to either conditions or exceptions
                    if row['is_exception'].lower() == 'true':
                        rules_dict[rule_name]['exceptions'][row['condition_type']].append(value)
                    else:
                        rules_dict[rule_name]['conditions'][row['condition_type']].append(value)

            # Convert dictionary to list of rules
            #rules = json.dumps(list(rules_dict.values()))

            # Convert to JSON object of rules
            rules = json.loads(json.dumps(rules_dict))
            ###*** something wrong here, maybe all should export json dumps than JSON object...???

            self.log_print(f"Successfully imported {len(rules)} rules from CSV")

            return rules  # Return a JSON object of the rules

        except Exception as e:
            self.log_print(f"Error importing rules from CSV: {str(e)}")
            self.log_print(f"Error details: {str(e.__class__.__name__)}")
            import traceback
            self.log_print(f"Traceback: {traceback.format_exc()}")
            return None

    def export_rules(self, rules_json=None, rules_file=OUTLOOK_RULES_FILE):
        """Export Outlook rules to CSV file"""
        try:
            if rules_json is None:   #this should never happen
                self.log_print("Rules JSON is Empty, do not overwrite rules_file CSV and exit with error")
                return None

            # Convert rules_json to JSON object if it's a string or dict
            if isinstance(rules_json, str):
                rules = json.loads(rules_json)
                self.log_print(f"export_rules: Found rules_json is a string and converted to JSON object")
            elif isinstance(rules_json, dict):
                rules = json.loads(json.dumps(rules_json))
                self.log_print(f"export_rules: Found rules_json is a dict and converted to JSON object")
            else:
                rules = rules_json
                self.log_print(f"export_rules: Found rules_json is a JSON object, no action taken")

            # Parse JSON string into Python object if it's a string
            rules = json.loads(rules_json) if isinstance(rules_json, str) else rules_json
            self.log_print(f"Processing {len(rules)} rules")

            # Read additional rules from an OUTLOOK_SAFE_SENDERS CSV file
            safe_senders = []
            if os.path.exists(OUTLOOK_SAFE_SENDERS_FILE):
                with open(OUTLOOK_SAFE_SENDERS_FILE, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line:
                            safe_senders.append(line)

            for rule in rules_json:
                if rule.get("name") == "SpamAutoDeleteBody":
                    if "body" not in rule["conditions"]:
                        rule["conditions"]["body"] = []
                    for sender in safe_senders:
                        rule["conditions"]["body"].append({"address": sender})

            with open(OUTLOOK_RULES_FILE, 'w', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=[
                    'last_modified',
                    'rule_name',
                    'enabled',
                    'execution_order',
                    'condition_type',
                    'condition_value',
                    'is_exception',
                    'action_delete',
                    'needs_special_handling'
                ])
                writer.writeheader()

                timestamp = datetime.now().isoformat()

                # Process each rule
                for rule in rules:
                    rule_name = rule['name']
                    enabled = rule['enabled']
                    execution_order = rule['executionOrder']
                    delete_action = rule['actions']['delete']

                    # Process conditions
                    for cond_type in ['from', 'subject', 'body', 'header']:
                        values = rule['conditions'].get(cond_type, [])
                        if values:  # Only process if there are values
                            for value in values:
                                if isinstance(value, dict) and 'address' in value:
                                    value = value['address']
                                escaped_value, needs_special = self._escape_pattern(str(value))
                                writer.writerow({
                                    'last_modified': timestamp,
                                    'rule_name': rule_name,
                                    'enabled': str(enabled).lower(),
                                    'execution_order': execution_order,
                                    'condition_type': cond_type,
                                    'condition_value': escaped_value,
                                    'is_exception': 'false',
                                    'action_delete': str(delete_action).lower(),
                                    'needs_special_handling': str(needs_special).lower()
                                })

                    # Process exceptions
                    for exc_type in ['from', 'subject', 'body', 'header']:
                        values = rule['exceptions'].get(exc_type, [])
                        if values:  # Only process if there are values
                            for value in values:
                                if isinstance(value, dict) and 'address' in value:
                                    value = value['address']
                                escaped_value, needs_special = self._escape_pattern(str(value))
                                writer.writerow({
                                    'last_modified': timestamp,
                                    'rule_name': rule_name,
                                    'enabled': str(enabled).lower(),
                                    'execution_order': execution_order,
                                    'condition_type': exc_type,
                                    'condition_value': escaped_value,
                                    'is_exception': 'true',
                                    'action_delete': str(delete_action).lower(),
                                    'needs_special_handling': str(needs_special).lower()
                                })

            self.log_print(f"Rules successfully exported to {self.rules_file}")
            return True

        except Exception as e:
            self.log_print(f"Error exporting rules: {str(e)}")
            self.log_print(f"Error details: {str(e.__class__.__name__)}")
            import traceback
            self.log_print(f"Traceback: {traceback.format_exc()}")
            return False


    def get_rules(self):
        """Get rules from CSV if available, otherwise from Outlook"""
        # disabling outlook rules via CSV file - was not working
        # CSV_rules = []
        # CSV_rules = self.import_rules()
        # self.log_print(f"Import rules from CSV ({OUTLOOK_RULES_FILE})")

        # outlook_rules = []
        self.log_print(f"Import rules from Outlook")
        outlook_rules = self.get_outlook_rules()

        # debugging - compare CSV_rules to Outlook_rules and print the differences between them
        # self.output_rules_differences(outlook_rules, CSV_rules) #outlook rules are currently the primary source

        # debugging - for this run, set the rules to be from Outlook
        rules = outlook_rules

        #To be moved elsewhere
        # self.log_print(f"Export rules to CSV ({OUTLOOK_RULES_FILE}): {rules}")
        # self.export_rules(rules)

        # debugging to show the rules
        # self.log_print(f"Rules loaded: {rules}")

        # add a check to convert to a JSON object (if it a string or dict)
        if isinstance(rules, str) or isinstance(rules, dict):
            rules = json.loads(json.dumps(rules))

        return rules

    def print_rules_summary(self, rules):   # rules should be a JSON object
        """Print a summary of all rules in the CSV file"""
        try:
            # add a check to convert to a JSON object (if it a string or dict)
            if isinstance(rules, str) or isinstance(rules, dict):
                rules = json.loads(json.dumps(rules))

            self.log_print("\nRules Summary:")
            for rule in rules:
                self.log_print(f"\nRule: {rule['name']} (Enabled: {rule['enabled']})")
                for cond_type, values in rule['conditions'].items():
                    if not isinstance(values, list):
                        values = [values]
                    self.log_print(f"  {cond_type} conditions:")
                    for value in values:
                        self.log_print(f"    - {value}")
                self.log_print("  Actions:")
                for action, value in rule['actions'].items():
                    self.log_print(f"    - {action}: {value}")

        except Exception as e:
            self.log_print(f"Error printing rules summary: {str(e)}")

    def combine_email_header_lines(self, email_header):
        """
        Combine email headers, handling lines split across multiple lines, and find the first line containing "from:".

        Args:
            email_headers (str): The email headers as a single string.

        Returns:
            str: The first line containing "from:", or None if not found.
        """
        # Build email_header, combining lines split across multiple lines into one line (combine From:)
        email_header_list = []
        for line in email_header.splitlines():
            if line.startswith((' ', '\t')):
                # Continuation line, append to the previous line
                email_header_list[-1] += ' ' + line.strip()
            else:
                # New header field
                email_header_list.append(line.strip())

        # Convert email_header_list back into a single string
        updated_email_header = '\n'.join(email_header_list)

        # Sanitize the updated email header
        updated_email_header = self._sanitize_string(updated_email_header)

        # Convert to lowercase for easier keyword matching
        updated_email_header = updated_email_header.lower()

        return updated_email_header

    def header_from(self, email_header):
        """
        Process email headers to find the first line containing "from:" and extract the domain.

        Args:
            email_header (str): The email headers as a single string.

        Returns:
            str: The domain extracted from the "from:" line, padded to 20 characters, or None if not found.
        """
        line_with_from = None

        # Iterate over each element in email_header
        for line in email_header.splitlines():
            if "from:" in line.lower():
                line_with_from = line
                break

        if line_with_from:
            from_domain = re.search(r'@[\w.-]+', line_with_from)
            if from_domain:
                from_domain_str = from_domain.group(0)
                return from_domain_str

        return None

    def from_report(self, emails_to_process, emails_added_info, rules_json):
        """
        Generate a report of emails with phishing indicators or no rule matches, including the From domain.

        Args:
            emails_to_process (list): List of emails to process.
            emails_added_info (list): List of dictionaries containing additional information about each email.
        """
        # Print a list for Phishing OR Match=false with From: "@<domain>.<>" so they can be easily added to the rules

        for email in emails_to_process:
            email_index = emails_to_process.index(email)
            try:
                if ("phishing_indicators" in emails_added_info[email_index] and
                    emails_added_info[email_index]["phishing_indicators"] is not None):
                    # Create a string from email.header for the From: line with format: "@<domain>.<> (20 characters or less,
                    # padded to 20) Email <n> (with 2 leading blanks)"


                    email_header = emails_added_info[email_index]["email_header"]
                    from_domain = self.header_from(email_header)

                    output_string = (from_domain.ljust(20) +
                                    f"| Email {email_index+1:>3} | " +
                                    f"Phishing indicators: {emails_added_info[email_index]['phishing_indicators']}")
                    self.log_print(f"{output_string}", level="INFO")
                    simple_print(f"{output_string}")
            except Exception as e:
                simple_print(f"Error processing phishing indicators for email: {str(e)}")

            try:
                if (emails_added_info[email_index]["match"] == False):
                    # Create a string from email.header for the From: line with format: "@<domain>.<> (20 characters or less,
                    # padded to 20) Email <n> (with 2 leading blanks)"

                    email_header = emails_added_info[email_index]["email_header"]
                    from_domain = self.header_from(email_header)

                    output_string = from_domain.ljust(20) + f"| Email {email_index+1:>3} | Matched no rules"
# 03/29/2025 Harold Kimmey Add functionality to update JSON rules by adding a simple Y/N to add to the "@<domain>"
# to the header rules in the JSON rules
                    self.log_print(f"{output_string}", level="INFO")
                    simple_print(f"{output_string}")

            except Exception as e:
                self.log_print(f"Error processing match = false email: {str(e)}")

    def get_unique_URL_stubs(self, email_body):
        unique_stubs = []
        seen_stubs = set()
        url_pattern = re.compile(r'(\.[\w-]+\.[\w-]+)|(/[\w-]+\.[\w-]+)')
        for line in email_body.splitlines():
            matches = url_pattern.findall(line)
            for match in matches:
                stub = match[0] if match[0] else match[1]
                # Remove leading "/" or "."
                cleaned_stub = stub.lstrip('/.')
                # Add both versions to the list if not seen before
                if '/' + cleaned_stub not in seen_stubs:
                    unique_stubs.append('/' + cleaned_stub)
                    seen_stubs.add('/' + cleaned_stub)
                if '.' + cleaned_stub not in seen_stubs:
                    unique_stubs.append('.' + cleaned_stub)
                    seen_stubs.add('.' + cleaned_stub)
        return unique_stubs

    def URL_report(self, emails_to_process, emails_added_info):
        """
        Generate a report of emails with phishing indicators or no rule matches,
            including unique URL stubs "/<domain>.<>" and ".<domain>.<>" from the body.

        Args:
            emails_to_process (list): List of emails to process.
            emails_added_info (list): List of dictionaries containing additional information about each email.
        """
        # Print a list for Phishing OR Match=false, report body unique URL stubs "/<domain>.<>" and ".<domain>.<>" so they can be easily added to the rules
        #     collect them all first, then determine uniqueness, then print one per line

        for email in emails_to_process:
            email_index = emails_to_process.index(email)
            try:
                if ("phishing_indicators" in emails_added_info[email_index] and
                    emails_added_info[email_index]["phishing_indicators"] is not None):
                    # Create a string from email.header for the From: line with format: "@<domain>.<> (20 characters or less,
                    # padded to 20) Email <n> (with 2 leading blanks)"

                    unique_URL_stubs = self.get_unique_URL_stubs(email.Body)

                    for stub in unique_URL_stubs:
                        output_string = (stub.ljust(30) +
                                    f"| Email {email_index+1:>3} | " +
                                    f"From: {self._sanitize_string(email.SenderEmailAddress)}")
                        self.log_print(f"{output_string}",level="INFO")
                        simple_print(f"{output_string}")
            except Exception as e:
                self.log_print(f"Error processing phishing indicators for email: {str(e)}")

            try:
                if (emails_added_info[email_index]["match"] == False):
                    # Create a string from email.header for the From: line with format: "@<domain>.<> (20 characters or less,
                    # padded to 20) Email <n> (with 2 leading blanks)"

                    unique_URL_stubs = self.get_unique_URL_stubs(email.Body)

                    for stub in unique_URL_stubs:
                        output_string = (stub.ljust(30) +
                                    f"| Email {email_index+1:>3} | " +
                                    f"From: {self._sanitize_string(email.SenderEmailAddress)}")
                        self.log_print(f"{output_string}",level="INFO")
                        simple_print(f"{output_string}")

            except Exception as e:
                self.log_print(f"Error processing match = false email: {str(e)}")

    def _process_conditions(self, conditions_obj, is_exception):
        """Helper method to process rule conditions or exceptions"""
        conditions = {}

        try:
            # From addresses
            if hasattr(conditions_obj, "From") and conditions_obj.From:
                try:
                    conditions["from"] = [
                        {
                            "address": recipient.Address if hasattr(recipient, "Address") else None,
                            "name": recipient.Name if hasattr(recipient, "Name") else None
                        }
                        for recipient in conditions_obj.From.Recipients
                    ]
                    # Print the contents of conditions["from"] #can be used for extra debugging information
                    if is_exception:
                        self.log_print(f"Exception conditions['from']: {conditions['from']}", "DEBUG")
                    else:
                        self.log_print(f"Conditions['from']: {conditions['from']}", "DEBUG")

                except Exception as e:
                    self.log_print(f"Error processing From condition: {str(e)}")
                    conditions["from"] = []

            # Subject keywords
            if hasattr(conditions_obj, "Subject") and conditions_obj.Subject:
                try:
                    if is_exception:
                        self.log_print(f"Exception conditions_obj.Subject.Text: {conditions_obj.Subject.Text}", "DEBUG")
                    else:
                        self.log_print(f"Conditions_obj.Subject.Text: {conditions_obj.Subject.Text}", "DEBUG")

                    if hasattr(conditions_obj.Subject, "Text"):
                        if isinstance(conditions_obj.Subject.Text, str):
                            subject_text = conditions_obj.Subject.Text
                        elif isinstance(conditions_obj.Subject.Text, tuple):
                            subject_text = "; ".join(conditions_obj.Subject.Text)
                        else:
                            subject_text = ""
                    else:
                        subject_text = ""
                    conditions["subject"] = [kw.strip() for kw in subject_text.split(";") if kw.strip()]
                except Exception as e:
                    self.log_print(f"Error processing Subject condition: {str(e)}")
                    conditions["subject"] = []

            # Body keywords
            if hasattr(conditions_obj, "Body") and conditions_obj.Body:
                try:
                    if is_exception:
                        self.log_print(f"Exception conditions_obj.Body.Text: {conditions_obj.Body.Text}", "DEBUG")
                    else:
                        self.log_print(f"Conditions_obj.Body.Text: {conditions_obj.Body.Text}", "DEBUG")

                    if hasattr(conditions_obj.Body, "Text"):
                        if isinstance(conditions_obj.Body.Text, str):
                            body_text = conditions_obj.Body.Text
                        elif isinstance(conditions_obj.Body.Text, tuple):
                            body_text = "; ".join(conditions_obj.Body.Text)
                            # self.log_print(f"body_text: {body_text}")
                        else:
                            body_text = ""
                    else:
                        body_text = ""
                    conditions["body"] = [kw.strip() for kw in body_text.split(";") if kw.strip()]
                except Exception as e:
                    self.log_print(f"Error processing Body condition: {str(e)}")
                    conditions["body"] = []

            # Header keywords
            if hasattr(conditions_obj, "MessageHeader") and conditions_obj.MessageHeader:
                try:
                    if is_exception:
                        self.log_print(f"Exception conditions_obj.MessageHeader.Text: {conditions_obj.MessageHeader.Text}", "DEBUG")
                    else:
                        self.log_print(f"Conditions_obj.MessageHeader.Text: {conditions_obj.MessageHeader.Text}", "DEBUG")

                    if hasattr(conditions_obj.MessageHeader, "Text"):
                        if isinstance(conditions_obj.MessageHeader.Text, str):
                            header_text = conditions_obj.MessageHeader.Text
                        elif isinstance(conditions_obj.MessageHeader.Text, tuple):
                            header_text = "; ".join(conditions_obj.MessageHeader.Text)
                        else:
                            header_text = ""
                    else:
                        header_text = ""
                    conditions["header"] = [kw.strip() for kw in header_text.split(";") if kw.strip()]
                except Exception as e:
                    self.log_print(f"Error processing Header condition: {str(e)}")
                    conditions["header"] = []

            # Attachment condition
            if hasattr(conditions_obj, "Attachment"):
                if is_exception:
                    self.log_print(f"Exception conditions_obj.Attachment: {bool(conditions_obj.Attachment)}", "DEBUG")
                else:
                    self.log_print(f"Conditions_obj.Attachment: {bool(conditions_obj.Attachment)}", "DEBUG")

                conditions["has_attachments"] = bool(conditions_obj.Attachment)

        except Exception as e:
            self.log_print(f"Error processing conditions: {str(e)}")
            conditions["error"] = str(e)

        return conditions

    def _process_actions(self, actions_obj):
        """Helper method to process rule actions"""
        actions = {}

        try:
            # Move to Folder
            if hasattr(actions_obj, "MoveToFolder") and actions_obj.MoveToFolder:
                try:
                    actions["move_to_folder"] = {
                        "folder_path": actions_obj.MoveToFolder.FolderPath if hasattr(actions_obj.MoveToFolder, "FolderPath") else None,
                        "folder_name": actions_obj.MoveToFolder.Name if hasattr(actions_obj.MoveToFolder, "Name") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing MoveToFolder action: {str(e)}")

            # Copy to Folder
            if hasattr(actions_obj, "CopyToFolder") and actions_obj.CopyToFolder:
                try:
                    actions["copy_to_folder"] = {
                        "folder_path": actions_obj.CopyToFolder.FolderPath if hasattr(actions_obj.CopyToFolder, "FolderPath") else None,
                        "folder_name": actions_obj.CopyToFolder.Name if hasattr(actions_obj.CopyToFolder, "Name") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing CopyToFolder action: {str(e)}")

            # Assign to Category
            if hasattr(actions_obj, "AssignToCategory") and actions_obj.AssignToCategory:
                try:
                    # Outlook may store one or more category names in a collection property.
                    # First, check if there is a Categories collection
                    if hasattr(actions_obj.AssignToCategory, "Categories") and actions_obj.AssignToCategory.Categories:
                        # Convert the collection into a list
                        category_collection = actions_obj.AssignToCategory.Categories
                        # Depending on the COM object, you might iterate over it
                        category_names = [cat for cat in category_collection]
                        # Join the names if more than one
                        category_name = ", ".join(category_names)
                    else:
                        # Fall back to a simple property "Category" if available.
                        category_name = getattr(actions_obj.AssignToCategory, "Category", None)
                    self.log_print(f"AssignToCategory action found, category_name: {category_name}")
                    actions["assign_to_category"] = {
                        "category_name": category_name
                    }
                except Exception as e:
                    self.log_print(f"Error processing AssignToCategory action: {str(e)}")
                    actions["assign_to_category"] = {
                        "category_name": None
                    }

            # Delete
            if hasattr(actions_obj, "Delete") and actions_obj.Delete:
                actions["delete"] = True

            # Stop processing more rules
            if hasattr(actions_obj, "StopProcessingMoreRules") and actions_obj.StopProcessingMoreRules:
                try:
                    self.log_print("StopProcessingMoreRules action found")
                    actions["stop_processing_more_rules"] = True
                except Exception as e:
                    self.log_print(f"Error processing StopProcessingMoreRules action: {str(e)}")

            # Mark as Read
            if hasattr(actions_obj, "MarkAsRead") and actions_obj.MarkAsRead:
                try:
                    self.log_print("MarkAsRead action found")
                    actions["mark_as_read"] = True
                except Exception as e:
                    self.log_print(f"Error processing MarkAsRead action: {str(e)}")

            # Clear the Message Flag
            if hasattr(actions_obj, "ClearFlag") and actions_obj.ClearFlag:
                try:
                    self.log_print("ClearFlag action found")
                    actions["clear_flag"] = True
                except Exception as e:
                    self.log_print(f"Error processing ClearFlag action: {str(e)}")

            # Forward
            if hasattr(actions_obj, "Forward") and actions_obj.Forward:
                try:
                    actions["forward"] = [
                        {
                            "address": recipient.Address if hasattr(recipient, "Address") else None,
                            "name": recipient.Name if hasattr(recipient, "Name") else None
                        }
                        for recipient in actions_obj.Forward.Recipients
                    ]
                except Exception as e:
                    self.log_print(f"Error processing Forward action: {str(e)}")
                    actions["forward"] = []

            # Redirect
            if hasattr(actions_obj, "Redirect") and actions_obj.Redirect:
                try:
                    actions["redirect"] = [
                        {
                            "address": recipient.Address if hasattr(recipient, "Address") else None,
                            "name": recipient.Name if hasattr(recipient, "Name") else None
                        }
                        for recipient in actions_obj.Redirect.Recipients
                    ]
                except Exception as e:
                    self.log_print(f"Error processing Redirect action: {str(e)}")
                    actions["redirect"] = []

            # Reply
            if hasattr(actions_obj, "Reply") and actions_obj.Reply:
                try:
                    actions["reply"] = {
                        "template": actions_obj.Reply.Template if hasattr(actions_obj.Reply, "Template") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing Reply action: {str(e)}")

            # Play Sound
            if hasattr(actions_obj, "PlaySound") and actions_obj.PlaySound:
                try:
                    actions["play_sound"] = {
                        "sound_file": actions_obj.PlaySound.SoundFile if hasattr(actions_obj.PlaySound, "SoundFile") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing PlaySound action: {str(e)}")

            # Display Desktop Alert
            if hasattr(actions_obj, "DisplayDesktopAlert") and actions_obj.DisplayDesktopAlert:
                actions["display_desktop_alert"] = True

            # Set Importance
            if hasattr(actions_obj, "SetImportance") and actions_obj.SetImportance:
                try:
                    actions["set_importance"] = {
                        "importance_level": actions_obj.SetImportance.ImportanceLevel if hasattr(actions_obj.SetImportance, "ImportanceLevel") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing SetImportance action: {str(e)}")

            # Set Sensitivity
            if hasattr(actions_obj, "SetSensitivity") and actions_obj.SetSensitivity:
                try:
                    actions["set_sensitivity"] = {
                        "sensitivity_level": actions_obj.SetSensitivity.SensitivityLevel if hasattr(actions_obj.SetSensitivity, "SensitivityLevel") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing SetSensitivity action: {str(e)}")

            # Print
            if hasattr(actions_obj, "Print") and actions_obj.Print:
                actions["print"] = True

            # Run Script
            if hasattr(actions_obj, "RunScript") and actions_obj.RunScript:
                try:
                    actions["run_script"] = {
                        "script_path": actions_obj.RunScript.ScriptPath if hasattr(actions_obj.RunScript, "ScriptPath") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing RunScript action: {str(e)}")

            # Start Application
            if hasattr(actions_obj, "StartApplication") and actions_obj.StartApplication:
                try:
                    actions["start_application"] = {
                        "application_path": actions_obj.StartApplication.ApplicationPath if hasattr(actions_obj.StartApplication, "ApplicationPath") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing StartApplication action: {str(e)}")

            # Mark as Task
            if hasattr(actions_obj, "MarkAsTask") and actions_obj.MarkAsTask:
                try:
                    actions["mark_as_task"] = {
                        "task_due_date": actions_obj.MarkAsTask.TaskDueDate if hasattr(actions_obj.MarkAsTask, "TaskDueDate") else None
                    }
                except Exception as e:
                    self.log_print(f"Error processing MarkAsTask action: {str(e)}")

        except Exception as e:
            self.log_print(f"Error processing actions: {str(e)}")

        return actions

    def check_phishing_indicators(self, email):
        """Check for phishing indicators in an email"""
        indicators = []

        try:
            # Check sender mismatch
            sender = email.SenderEmailAddress.lower()
            display_name = email.SenderName.lower()
            if '@' in display_name and display_name != sender:
                self.log_print(f"Phishing indicator: Sender name/email mismatch: {display_name} vs {sender}")
                indicators.append("Phishing indicator: Sender name/email mismatch")

            # Check urgent language
            urgent_words = ['urgent', 'immediate', 'action required', 'account suspended']
            found_urgent = [word for word in urgent_words if word in email.Subject.lower()]
            if found_urgent:
                self.log_print(f"Phishing indicator: Found urgent language in subject: {found_urgent}")
                indicators.append("Phishing indicator: Found urgent language in subject")

            # Check URLs
            if email.HTMLBody:
                href_pattern = r'href=[\'"]?([^\'" >]+)'
                urls = re.findall(href_pattern, email.HTMLBody)
                for url in urls:
                    if 'http' in url.lower():
                        if url.lower() not in email.HTMLBody.lower():
                            self.log_print(f"Phishing indicator: Found mismatched URL display text: {url}")
                            indicators.append("Phishing indicator: Found Mismatched URL display text")
                            break

            # Check sensitive words
            sensitive_words = ['password', 'login', 'credential', 'verify account']
            found_sensitive = [word for word in sensitive_words if word in email.Body.lower()]
            if found_sensitive:
                self.log_print(f"Phishing indicator: Found requests for sensitive information: {found_sensitive}")
                indicators.append("Phishing indicator: Found requests for sensitive information")

        except Exception as e:
            self.log_print(f"Error checking indicators: {str(e)}")

        return indicators

    def delete_email_with_retry(self, email, max_retries=10, delay=1):
        """
        Attempt to delete an email with retries.

        Args:
            email: The email object to delete.
            max_retries: Maximum number of retries.
            delay: Delay between retries in seconds.
        """

        import time
        for attempt in range(max_retries):
            try:
                email.Delete()
                self.log_print(f"Email deleted successfully on attempt {attempt + 1}")
                return
            except Exception as e:
                self.log_print(f"Error deleting email on attempt {attempt + 1}: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise
        return

    def mark_email_read_with_retry(self, email, max_retries=10, delay=1):
        """
        Attempt to mark an email as unread with retries.

        Args:
            email: The email object to mark as unread.
        """

        import time
        for attempt in range(max_retries):
            try:
                if email.UnRead:
                    email.UnRead = False
                    email.Save()
                    self.log_print(f"Email marked as read successfully on attempt  {attempt + 1}")
                return
            except Exception as e:
                self.log_print(f"Error marking email as read on attempt {attempt + 1}: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise
        return

    def clear_email_flag_with_retry(self, email, max_retries=10, delay=1):
        """
        Attempt to clear the flag on an email; with with retries.

        Args:
            email: The email object to clear the flag.
        """

        import time
        for attempt in range(max_retries):
            try:
                email.Flag.Clear()
                # email.Save()
                self.log_print(f"Email flag cleared successfully on attempt  {attempt + 1}")
                return
            except Exception as e:
                self.log_print(f"Error clearing flag on email on attempt {attempt + 1}: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise
        return

    def assign_category_to_email_with_retry(self, email, category_name, max_retries=10, delay=1):
        """
        Attempt to mark an email as unread with retries.

        Args:
            email: The email object to mark as unread.
        """

        import time
        for attempt in range(max_retries):
            try:
                email.Categories = category_name
                email.Save()
                self.log_print(f"Email category {category_name} assigned successfully on attempt  {attempt + 1}")
                return
            except Exception as e:
                self.log_print(f"Error assigning {category_name} to email on attempt {attempt + 1}: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise
        return

    def process_emails(self, rules_json, days_back=DAYS_BACK_DEFAULT):
        """Process emails based on the rules in the rules_json object"""
        self.log_print(f"\n\nStarting email processing")
        self.log_print(f"Target folder: {self.target_folder.Name}", "DEBUG")
        self.log_print(f"Processing emails from last {days_back} days")

        try:
            # Parse the rules from the JSON object
            rules = json.loads(rules_json) if isinstance(rules_json, str) else rules_json
            # Ensure rules is a list of rule objects
            rules = [rules_json] if isinstance(rules_json, dict) else rules_json

            # Get recent emails from the target folder
            restriction = "[ReceivedTime] >= '" + \
                (datetime.now() - timedelta(days=days_back)).strftime('%m/%d/%Y') + "'"
            emails = self.target_folder.Items.Restrict(restriction)

            if not emails:
                self.log_print("No emails found to process.")

            if isinstance(emails, str):
                self.log_print("Error: 'emails' is a string, expected a collection of email objects.")

            emails.Sort("[ReceivedTime]", Descending=True)
            self.log_print(f"Total emails found: {emails.Count}")

            processed_count = 0
            flagged_count = 0
            deleted_total = 0
            matched_emails = []
            non_matched_emails = []

            self.log_print("Beginning email analysis:")

            # Create a list of emails to process (done because if deleting emails in "email in emails") it will skip emails
            emails_to_process = [email for email in emails]
            self.log_print(f"before adding fields to emails_added_info")
            emails_added_info = [{
                "match": False,
                "rule": "",
                "matched_keyword": "",
                "indicators": [],
                "email_header": ""
            } for email in emails_to_process]
            self.log_print(f"after adding fields to emails_added_info")

            for email in emails_to_process:
                try:
                    processed_count += 1
                    email_index = emails_to_process.index(email)
                    email_deleted = False
                    if (DEBUG) and (processed_count > DEBUG_EMAILS_TO_PROCESS):
                        self.log_print(f"Debug mode: Stopping after {DEBUG_EMAILS_TO_PROCESS} emails")
                        return
                    email_header = self.combine_email_header_lines(email.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E"))
                    self.log_print(f"\n\nEmail {processed_count}:")
                    self.log_print(f"Subject: {self._sanitize_string(email.Subject)}")
                    self.log_print(f"From: {self._sanitize_string(email.SenderEmailAddress)}")
                    self.log_print(f"Received: {email.ReceivedTime}")

                    # Sort rules to ensure delete actions are processed last
                    rules.sort(key=lambda rule: rule['actions'].get('delete', False))

                    for rule in rules:
                        if not isinstance(rule, dict) or 'actions' not in rule:
                            self.log_print(f"Invalid rule format: {rule}")
                            continue
                        if email_deleted:
                            continue  # Go to the next email if one rule deletes the current email
                        conditions = rule['conditions']
                        exceptions = rule['exceptions']
                        # print(rule, conditions) #can be used for extra debugging information
                        match = False

                        # Check 'from' addresses
                        if 'from' in conditions:
                            from_addresses = [addr['address'].lower() for addr in conditions['from']]
                            if any(addr in email.SenderEmailAddress.lower() for addr in from_addresses):
                                match = True
                                matched_keyword = next((addr['address'] for addr in conditions['from'] if addr['address'].lower() in email.SenderEmailAddress.lower()), None)
                                self.log_print(f"Matched keyword in from address: {matched_keyword}")
                                self.log_print(f"From: {self._sanitize_string(email.SenderEmailAddress)}")

                        # Check subject keywords
                        if 'subject' in conditions:
                            if any(keyword.lower() in email.Subject.lower() for keyword in conditions['subject']):
                                match = True
                                matched_keyword = next((keyword for keyword in conditions['subject'] if keyword.lower() in email.Subject.lower()), None)
                                self.log_print(f"Matched keyword in subject: {matched_keyword}")
                                self.log_print(f"Subject: {self._sanitize_string(email.Subject)}")

                        # Check body keywords
                        if 'body' in conditions:
                            if any(keyword.lower() in email.Body.lower() for keyword in conditions['body']):
                                match = True
                                matched_keyword = next((keyword for keyword in conditions['body'] if keyword.lower() in email.Body.lower()), None)
                                self.log_print(f"Matched keyword in body: {matched_keyword}")
                                matched_lines = [line for line in email.Body.splitlines() if matched_keyword.lower() in line.lower()]
                                if matched_lines:
                                    self.log_print(f"First line of body that matches the keyword: {matched_lines[0]}")
                                # below will print all the body lines that match if needed for debugging
                                if DEBUG:
                                    for line in email.Body.splitlines():
                                        if any(keyword.lower() in line.lower() for keyword in conditions['body']):
                                            self.log_print(f"Body: {line}", "DEBUG")
                        # Check header keywords
                        if 'header' in conditions:
                            if any(keyword.lower() in email_header for keyword in conditions['header']):
                                match = True
                                matched_keyword = next((keyword for keyword in conditions['header'] if keyword.lower() in email_header.lower()), None)
                                self.log_print(f"Matched keyword in header: {matched_keyword}")
                                matched_lines = [line for line in email_header.splitlines() if matched_keyword.lower() in line.lower()]
                                if matched_lines:
                                    self.log_print(f"First line of header that matches the keyword: {matched_lines[0]}")
                                # below will print all the body lines that match if needed for debugging
                                # for header in email_header.splitlines():
                                if DEBUG:
                                    for header in email_header.splitlines():
                                        if any(keyword.lower() in header.lower() for keyword in conditions['header']):
                                            self.log_print(f"Header: {header}", "DEBUG")

                        # Check for attachments - not using. could be added later - will need to be updated; will not work as-is
                        # if 'has_attachments' in conditions:
                        #     if bool(email.Attachments.Count > 0) != conditions['has_attachments']:
                        #         match = True

                    # Check exceptions
                        if 'from' in exceptions:
                            from_addresses = [addr['address'].lower() for addr in exceptions['from']]
                            if any(addr in email.SenderEmailAddress.lower() for addr in from_addresses):
                                match = False
                                matched_keyword = next((addr['address'] for addr in exceptions['from'] if addr['address'].lower() in email.SenderEmailAddress.lower()), None)
                                self.log_print(f"Exception matched keyword in from address: {matched_keyword}")
                                self.log_print(f"From: {self._sanitize_string(email.SenderEmailAddress)}")

                        # Check subject keywords in exceptions
                        if 'subject' in exceptions:
                            if any(keyword.lower() in email.Subject.lower() for keyword in exceptions['subject']):
                                match = False
                                matched_keyword = next((keyword for keyword in exceptions['subject'] if keyword.lower() in email.Subject.lower()), None)
                                self.log_print(f"Exception matched keyword in subject: {matched_keyword}")
                                self.log_print(f"Subject: {self._sanitize_string(email.Subject)}")

                        # Check body keywords in exceptions
                        if 'body' in exceptions:
                            if any(keyword.lower() in email.Body.lower() for keyword in exceptions['body']):
                                match = False
                                matched_keyword = next((keyword for keyword in exceptions['body'] if keyword.lower() in email.Body.lower()), None)
                                self.log_print(f"Exception matched keyword in body: {matched_keyword}")
                                self.log_print(f"Body: {self._sanitize_string(email.Body)}")

                        # Check header keywords in exceptions
                        if 'header' in exceptions:
                            if any(keyword.lower() in email_header for keyword in exceptions['header']):
                                match = False
                                matched_keyword = next((keyword for keyword in exceptions['header'] if keyword.lower() in email_header.lower()), None)
                                self.log_print(f"Exception matched keyword in header: {matched_keyword}")
                                for header in email_header.splitlines():
                                    self.log_print(f"Header: {header}")

                        # # Check for attachments - not using. could be added later - will need to be updated; will not work as-is
                        # if 'has_attachments' in conditions:
                        #     if bool(email.Attachments.Count > 0) != conditions['has_attachments']:
                        #         match = False

                        # If match is true need to process 2 things, but do them in separate steps
                        # first, if matched save in the copy of emails, add the rule and the keyword matched
                        #   If not matched, will pull the information from the email
                        # can use the original email if the index is available - can use the index of emails_added_info
                        if match:
                            emails_added_info[email_index]["match"] = match
                            emails_added_info[email_index]["rule"] = rule
                            emails_added_info[email_index]["matched_keyword"] = matched_keyword
                            emails_added_info[email_index]["email_header"] = email_header
                        else:
                            emails_added_info[email_index]["match"] = match
                            emails_added_info[email_index]["rule"] = None
                            emails_added_info[email_index]["matched_keyword"] = ""
                            emails_added_info[email_index]["email_header"] = email_header

                        if match:
                            self.log_print(f"Email matches rule: {rule['name']}")
                            # Perform actions based on the rule
                            actions = rule['actions']
                            self.log_print(f"Performing actions: {actions}")
                            # # removed the following if block as it now should be in rule JSON object
                            # if 'assign_to_category' in actions and actions['assign_to_category']['category_name']:
                            #     rule_name = rule['name']

                            #     category_name = self.rule_to_category.get(rule_name, actions['assign_to_category']['category_name'])
                            #     email.Categories = category_name
                            #     email.Save()
                            # self.log_print(f"Email assigned to category '{category_name}'")
                            if 'assign_to_category' in actions and actions['assign_to_category']['category_name']:
                                try: # to assign category based on rule name
                                    category_name = actions['assign_to_category']['category_name']
                                    self.assign_category_to_email_with_retry(email, category_name)
                                    self.log_print(f"Email assigned to category '{category_name}'", "DEBUG")
                                except Exception as e:
                                    self.log_print(f"Error assigning category to email: {str(e)}")
                                # # ***category name is now being added to the rules JSON object during get_outlook_rules
                                # email.Categories = actions['assign_to_category']['category_name']
                                # email.Save()
                                # self.log_print(f"Email assigned to category '{actions['assign_to_category']['category_name']}'")
                            if 'mark_as_read' in actions and actions['mark_as_read']:
                                # this flag is not being passed by outlook, so will never be set.  Keeping in case fixed in the future
                                if email.UnRead:
                                    self.mark_email_read_with_retry(email)
                                    self.log_print("Email marked as read")
                            if 'clear_flag' in actions and actions['clear_flag']:
                                # this flag is not being passed by outlook, so will never be set.  Keeping in case fixed in the future
                                self.clear_email_flag_with_retry(email)
                                self.log_print("Email flag cleared")
                            if 'set_importance' in actions and actions['set_importance']['importance_level']:
                                email.Importance = actions['set_importance']['importance_level']
                                email.Save()
                                self.log_print(f"Email importance set to {actions['set_importance']['importance_level']}")
                            if 'set_sensitivity' in actions and actions['set_sensitivity']['sensitivity_level']:
                                email.Sensitivity = actions['set_sensitivity']['sensitivity_level']
                                email.Save()
                                self.log_print(f"Email sensitivity set to {actions['set_sensitivity']['sensitivity_level']}")
                            if 'mark_as_task' in actions and actions['mark_as_task']['task_due_date']:
                                email.TaskDueDate = actions['mark_as_task']['task_due_date']
                                email.Save()
                                self.log_print(f"Email marked as task with due date: {actions['mark_as_task']['task_due_date']}")
                            if 'play_sound' in actions and actions['play_sound']['sound_file']:
                                import winsound
                                winsound.PlaySound(actions['play_sound']['sound_file'], winsound.SND_FILENAME)
                                self.log_print(f"Played sound: {actions['play_sound']['sound_file']}")
                            if 'display_desktop_alert' in actions and actions['display_desktop_alert']:
                                self.log_print("Desktop alert displayed")
                                # Implement desktop alert display logic here
                            if 'copy_to_folder' in actions and actions['copy_to_folder']['folder_name']:
                                folder_name = actions['copy_to_folder']['folder_name']
                                target_folder = self.target_folder.Folders[folder_name]
                                email.Copy().Move(target_folder)
                                self.log_print(f"Email copied to '{folder_name}' folder")
                            if 'forward' in actions and actions['forward']:
                                forward_recipients = [recipient['address'] for recipient in actions['forward']]
                                forward_email = email.Forward()
                                forward_email.To = ";".join(forward_recipients)
                                forward_email.Send()
                                self.log_print(f"Email forwarded to: {', '.join(forward_recipients)}")
                            if 'reply' in actions and actions['reply']['template']:
                                reply_email = email.Reply()
                                reply_email.Body = actions['reply']['template']
                                reply_email.Send()
                                self.log_print("Auto-reply sent")
                            if 'redirect' in actions and actions['redirect']:
                                redirect_recipients = [recipient['address'] for recipient in actions['redirect']]
                                redirect_email = email.Forward()
                                redirect_email.To = ";".join(redirect_recipients)
                                redirect_email.Send()
                                self.log_print(f"Email redirected to: {', '.join(redirect_recipients)}")
                            if 'print' in actions and actions['print']:
                                email.PrintOut()
                                self.log_print("Email printed")
                            if 'run_script' in actions and actions['run_script']['script_path']:
                                exec(open(actions['run_script']['script_path']).read())
                                self.log_print(f"Script executed: {actions['run_script']['script_path']}")
                            if 'start_application' in actions and actions['start_application']['application_path']:
                                import subprocess
                                subprocess.Popen(actions['start_application']['application_path'])
                                self.log_print(f"Application started: {actions['start_application']['application_path']}")
                            if 'move_to_folder' in actions and actions['move_to_folder']['folder_name']:
                                folder_name = actions['move_to_folder']['folder_name']
                                target_folder = self.target_folder.Folders[folder_name]
                                email.Move(target_folder)
                                self.log_print(f"Email moved to '{folder_name}' folder")
                            if 'stop_processing_more_rules' in actions and actions['stop_processing_more_rules']:
                                self.log_print("Stopping processing more rules")
                                # this flag is not being passed by outlook, so will never be set.  Keeping in case fixed in the future
                            if 'delete' in actions and actions['delete']:
                                try: #to mark email as read if unread
                                    if email.UnRead:
                                        self.mark_email_read_with_retry(email)
                                        # email.UnRead = False  # Delete implies marking the item as read
                                        self.log_print(f"Email marked as read", "DEBUG")
                                except:
                                    self.log_print(f"Error marking email as read", "DEBUG")

                                try: #to clear the flag on email
                                    if hasattr(email, 'Flag'):
                                        self.clear_email_flag_with_retry(email)
                                        # email.Flag.Clear()      # Delete implies clearing the flag
                                        self.log_print(f"Email flag was cleared", "DEBUG")
                                except:
                                    self.log_print(f"Error clearing flag", "DEBUG")

                                # Now always done in category assignment above and no longer needed here.
                                # try: # to assign category based on rule name
                                #     rule_name = rule['name']
                                #     category_name = self.rule_to_category.get(rule_name, actions['assign_to_category']['category_name'])
                                #     self.assign_category_to_email_with_retry(email, category_name)
                                #     # email.Categories = category_name
                                #     # email.Save()
                                #     self.log_print(f"Email assigned to category '{category_name}'", "DEBUG")
                                # except Exception as e:
                                #     self.log_print(f"Error assigning category to email: {str(e)}")

                                try:
                                    # delete email
                                    self.delete_email_with_retry(email)
                                    email_deleted = True
                                    deleted_total += 1
                                    self.log_print("Email marked as read, flag cleared and deleted")
                                    # self.simple_print(f"Deleted email from: {self._sanitize_string(email.SenderEmailAddress)}")
                                    # delete implies "Stop Processing More Rules".  Continue will go to next email
                                except Exception as e:
                                    self.log_print(f"Error deleting email: {str(e)}")

                                break # If delete, then process no more rules and go to next email
                            continue  # Go to the next email if one rule matches

                    # After all email rules are processed and it did not match any rules and the email has not been deleted, then check for phishing indicators
                    if not (email_deleted):
                        indicators = self.check_phishing_indicators(email)
                        if indicators:
                            flagged_count += 1
                            self.log_print(f"Phishing indicators found: {indicators}")
                            emails_added_info[email_index]["phishing_indicators"] = indicators
                        else:
                            self.log_print("No conditions or phishing indicators found")
                        # If it is in the Bulk Mail folder, but nothing indicated via rules or phishing,
                        # show the body and header, so we information needed to add it to a rule
                        for line in email.Body.splitlines():
                            self.log_print(f"Body: {line}")
                        for header in email_header.splitlines():
                            self.log_print(f"Header: {header}")

                except Exception as e:
                    self.log_print(f"Error processing email: {str(e)}")

            #****
            # Process a list of deleted emails with a one line summary of each via simple_print
            #     create a function deleted_report(emails_to_process, emails_added_info) to process the list

            # for Match=false, report header "<subject>  " so they can be easily added to the rules
            #     create a function header_report(emails_to_process, emails_added_info) to process the list

            # Print a list for Phishing OR Match=false, report body unique URL stubs "/<domain>.<>" and ".<domain>.<>" so they can be easily added to the rules
            #     collect them all first, then determine uniqueness, then print one per line
            self.log_print(f"\nProcessing Report of URL's from phishing or match = False")
            self.URL_report(emails_to_process, emails_added_info)

            # Print a list for Phishing OR Match=false with From: "@<domain>.<>" so they can be easily added to the rules
            self.log_print(f"\nProcessing Report of From's from phishing or match = False")
            self.from_report(emails_to_process, emails_added_info, rules_json)

            self.log_print(f"\nProcessing Summary:")
            self.log_print(f"Processed {processed_count} emails")
            self.log_print(f"Flagged {flagged_count} emails as possible Phishing attempts")
            self.log_print(f"Deleted {deleted_total} emails")
            self.log_print(f"END of Run =============================================================\n\n")

            simple_print(f"\nProcessing Summary:")
            simple_print(f"Processed {processed_count} emails")
            simple_print(f"Flagged {flagged_count} emails as possible Phishing attempts")
            simple_print(f"Deleted {deleted_total} emails")

        except Exception as e:
            self.log_print(f"Error in process_emails: {str(e)}")
            raise



# Main program execution
def main():
    """Main function to run the security agent"""

    # Initialize agent with debug mode enabled
    agent = OutlookSecurityAgent()  # call with defaults

    try:
        simple_print(f"\n=============================================================")
        simple_print(f"Starting Outlook Security Agent at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}")
        simple_print(f"This will make changes")
        simple_print(f"Check the {OUTLOOK_SECURITY_LOG} for detailed information")
        agent.log_print(f"\n=============================================================")
        agent.log_print(f"Starting Outlook Security Agent at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}")
        agent.log_print(f"This will make changes")
        agent.log_print(f"Check the {OUTLOOK_SECURITY_LOG} for detailed information")

        rules_json = agent.get_rules()  # updated for new CSV code - was get_outlook_rules()
        rules_before = rules_json
        #simple_print(f"JSON Rules\n{rules_json}") if DEBUG else None

        # Process last N days of emails - see DAYS_BACK_DEFAULT
        agent.process_emails(rules_json)

        # Export rules if they've been updated
        # if rules_before != rules_json:
        #     agent.export_rules()

        export_rules_yaml(rules_json, YAML_RULES_ALL_FILE)

        simple_print(f"Execution complete at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}. Check the log file for detailed analysis:\n{OUTLOOK_SECURITY_LOG}")
        simple_print(f"=============================================================\n")
        agent.log_print(f"Execution complete at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}. Check the log file for detailed analysis:\n{OUTLOOK_SECURITY_LOG}")
        agent.log_print(f"============================================================={CRLF}{CRLF}")

    except Exception as e:
        simple_print(f"\nError: {str(e)}")
        logging.error(f"Main execution error: {str(e)}")

if __name__ == "__main__":
    main()
