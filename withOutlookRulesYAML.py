# 01/24/2025 Harold Kimmey All working as expected
# 02/10/2025 Harold Kimmey Completed move to www.github.com/kimmeyh/spamfilter.git repository
# 02/17/2025 Harold Kimmey Updated _process_actions to accurately pull the assign_to_category value by searching it as an object
# 03/28/2025 Harold Kimmey Exported rules to JSON so that they can be maintained in a separate YAML file (the can be transferred between machines and platforms (Windows, Mac, Linux, Android, iOS, etc.))
#   Spam filter rules - done
#   Safe Senders - done
#   Safe recipients - done (was empty)
#   Blocked Senders - very small and were added manually to Outlook Rules - spam body# 03/30/2025 Harold Kimmey Added export and import of YAML rules
# 04/01/2025 Harold Kimmey Verified export of rules from Outlook to YAML file (at exit) matches rules from import of YAML file
# 04/01/2025 Harold Kimmey Switch to using YAML file as import instead of Outlook rules
# 04/01/2025 Harold Kimmey Committed changes, pushed, PR to Main branch of kimmeyh/spamfilter.git
# 05/13/2025 Harold Kimmey Current Status
#   - All rules are being read from the YAML file
#   - All rules are being written to the YAML file at the end of the run successfully
#   - All safe_sender rules are being read from the YAML file
#   - All safe_sender rules are being written to the YAML file (updated)
#   - Removed checks for updates to rules and safe_senders - instead will save copies to Archive for each run
#   - Update Output of rules.yaml - ensure they are written as compatible with regex strings  (double quoted, sorted, no duplicates)
#   - Output of safe_senders.yaml - ensure they are written as compatible with regex strings (double quoted), sorted, and unique
#   - Updated Output of rules.yaml - ensure each rule is sorted, and unique
# 05/19/2025 Harold Kimmey - Updates for feature/userinputheader
# - Need to add Next ****
#       Move backup files to a "backup directory"
#       Update mail processing to use safe_senders list for all header exceptions


#------------------List of future enhancements------------------
# Where is the best place to add updates to rules based on emails not deleted
# Add updates to rules for emails not deleted
#   for each email not deleted
#      show details of the email:  subject, from in header, URL's in the body
#       Suggest to add new domains (based on from in header) to the header rules
#       If N to header rule, suggest body rules
#       If no body rules added, suggest subject rules
#       Full commit after each of the above changes


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
# Successfully export rules to a yaml file that logically matches the JSON object at the end of the run
# Successfully import rules from the yaml file at the beginning of the run that matches the JSON object from get_outlook_rules
# Start to use the yaml file as the primary source of rules
# Add logic to add items to the rules JSON object via user input at the end of the run
#   - verify that the yaml file is updated with the new rules and can read them back in successfully
# Move all the appropriate rules to a yaml file structure
#   read from yaml and convert back to JSON object
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

#------------------General Documentation------------------
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

# Key Notes:
# - The YAML_RULES_FILE and YAML_RULES_SAFE_SENDERS_FILE
#   - all strings are maintained during export as trim(lowercase(str)))
#   - all lists of strings are sorted and unique (no duplicates)



#Imports for python base packages
import re
from datetime import datetime, timedelta
import logging
import sys
import json
import os
import yaml
import copy
import traceback

#Imports for packages that need to be installed
import win32com.client
import IPython

# Settings:
DEBUG = True # True or False
INFO = False if DEBUG else True #If not debugging, then INFO level logging
DEBUG_EMAILS_TO_PROCESS = 100 #100 for testing

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
YAML_RULES_PATH = f"D:/data/harold/github/OutlookMailSpamFilter/"
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
OUTLOOK_RULES_SUBSET = "SpamAutoDelete"
DAYS_BACK_DEFAULT = 365 # default number of days to go back in the calendar
CRLF = "\n"             # Carriage return and line feed for formatting


def simple_print(message):
    """Print message to a file or stdout based on OUTLOOK_SIMPLE_LOG"""
    if OUTLOOK_SIMPLE_LOG:
        with open(OUTLOOK_SIMPLE_LOG, 'a') as f:
            f.write(message + '\n')
    else: #write to the console
        print(message)

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
        self.YAMO_RULES_PATH = YAML_RULES_PATH  # Set appropriate default path
        self.YAML_RULES_FILE = YAML_RULES_FILE  # Set appropriate default file name
        self.YAML_SAFE_SENDERS_FILE = YAML_RULES_SAFE_SENDERS_FILE  # Set appropriate default file name

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
        self.log_print(f"\n=============================================================\nStarting new run")
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
        except Exception as e:
            self.log_print(f"Error: {str(e)}")
        return

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

    def _deep_compare_dicts(self, dict1, dict2, path=""):
        """Recursively compare dictionaries and return specific differences."""
        differences = []

        if not isinstance(dict1, dict) or not isinstance(dict2, dict):
            if dict1 != dict2:
                differences.append({
                    'path': path,
                    'value1': dict1,
                    'value2': dict2
                })
            return differences

        # Keys in dict1 but not in dict2
        for key in dict1:
            if key not in dict2:
                differences.append({
                    'path': f"{path}.{key}" if path else key,
                    'value1': dict1[key],
                    'value2': None
                })
                continue

            # If both have the key, compare values
            if isinstance(dict1[key], dict) and isinstance(dict2[key], dict):
                # Recursive comparison for nested dictionaries
                nested_diffs = self._deep_compare_dicts(
                    dict1[key], dict2[key],
                    f"{path}.{key}" if path else key
                )
                differences.extend(nested_diffs)
            elif isinstance(dict1[key], list) and isinstance(dict2[key], list):
                # Compare lists item by item
                list_diffs = self._deep_compare_lists(
                    dict1[key], dict2[key],
                    f"{path}.{key}" if path else key
                )
                differences.extend(list_diffs)
            elif dict1[key] != dict2[key]:
                differences.append({
                    'path': f"{path}.{key}" if path else key,
                    'value1': dict1[key],
                    'value2': dict2[key]
                })

        # Keys in dict2 but not in dict1
        for key in dict2:
            if key not in dict1:
                differences.append({
                    'path': f"{path}.{key}" if path else key,
                    'value1': None,
                    'value2': dict2[key]
                })

        return differences

    def _deep_compare_lists(self, list1, list2, path=""):
        """Compare two lists and return differences."""
        differences = []

        # Check for length differences
        if len(list1) != len(list2):
            differences.append({
                'path': path,
                'value1': f"List length: {len(list1)}",
                'value2': f"List length: {len(list2)}"
            })

        # Compare elements
        for i, (item1, item2) in enumerate(zip(list1, list2)):
            if isinstance(item1, dict) and isinstance(item2, dict):
                nested_diffs = self._deep_compare_dicts(
                    item1, item2,
                    f"{path}[{i}]"
                )
                differences.extend(nested_diffs)
            elif isinstance(item1, list) and isinstance(item2, list):
                nested_diffs = self._deep_compare_lists(
                    item1, item2,
                    f"{path}[{i}]"
                )
                differences.extend(nested_diffs)
            elif item1 != item2:
                differences.append({
                    'path': f"{path}[{i}]",
                    'value1': item1,
                    'value2': item2
                })

        # Handle different length lists
        for i in range(min(len(list1), len(list2)), max(len(list1), len(list2))):
            if i < len(list1):
                differences.append({
                    'path': f"{path}[{i}]",
                    'value1': list1[i],
                    'value2': "Missing"
                })
            else:
                differences.append({
                    'path': f"{path}[{i}]",
                    'value1': "Missing",
                    'value2': list2[i]
                })

        return differences

    def compare_rules(self, rules1, rules2):
        """Compare two sets of rules and return the differences."""
        # Extract rules arrays if wrapped in dictionaries
        if isinstance(rules1, dict) and "rules" in rules1:
            rules1_list = rules1["rules"]
        else:
            rules1_list = [rules1] if isinstance(rules1, dict) else rules1

        if isinstance(rules2, dict) and "rules" in rules2:
            rules2_list = rules2["rules"]
        else:
            rules2_list = [rules2] if isinstance(rules2, dict) else rules2

        # Validate input data types
        if not isinstance(rules1_list, list):
            self.log_print(f"Error: First rule set is not a list or dict. Type: {type(rules1_list)}")
            rules1_list = []
        if not isinstance(rules2_list, list):
            self.log_print(f"Error: Second rule set is not a list or dict. Type: {type(rules2_list)}")
            rules2_list = []

        # Create dictionaries keyed by rule name for easy comparison
        rules1_dict = {}
        rules2_dict = {}

        # Safely create dictionaries with error handling
        for i, rule in enumerate(rules1_list):
            if isinstance(rule, dict) and 'name' in rule:
                rules1_dict[rule['name']] = rule
            else:
                self.log_print(f"Warning: Invalid rule format in first set at index {i}: {type(rule)} - {rule}")
                # Optionally add more detailed debugging
                if isinstance(rule, dict):
                    self.log_print(f"Dictionary keys: {list(rule.keys())}")

        for i, rule in enumerate(rules2_list):
            if isinstance(rule, dict) and 'name' in rule:
                rules2_dict[rule['name']] = rule
            else:
                self.log_print(f"Warning: Invalid rule format in second set at index {i}: {type(rule)} - {rule}")
                # Optionally add more detailed debugging
                if isinstance(rule, dict):
                    self.log_print(f"Dictionary keys: {list(rule.keys())}")

        # Find rules unique to each set
        rules_only_in_1 = set(rules1_dict.keys()) - set(rules2_dict.keys())
        rules_only_in_2 = set(rules2_dict.keys()) - set(rules1_dict.keys())

        # Find modified rules (present in both but different)
        modified_rules = {}
        common_rules = set(rules1_dict.keys()) & set(rules2_dict.keys())
        for rule_name in common_rules:
            # Use deep comparison to identify specific differences
            diffs = self._deep_compare_dicts(rules1_dict[rule_name], rules2_dict[rule_name])
            if diffs:
                modified_rules[rule_name] = {
                    'rules1': rules1_dict[rule_name],
                    'rules2': rules2_dict[rule_name],
                    'differences': diffs
                }

        return {
            'rules_only_in_1': [rules1_dict[name] for name in rules_only_in_1],
            'rules_only_in_2': [rules2_dict[name] for name in rules_only_in_2],
            'modified_rules': modified_rules
        }


    def output_rules_differences(self, rule_set_one, rule_set_one_name, rule_set_two, rule_set_two_name):
        """Output the differences between 2 sets of JSON rules"""
        differences = self.compare_rules(rule_set_one, rule_set_two)

        # Print the differences
        self.log_print(f"{CRLF}Differences between Outlook rules and YAML rules:")
        if differences['rules_only_in_1']:
            self.log_print(f"\nRules only in {rule_set_one_name}:")
            for rule in differences['rules_only_in_1']:
                self.log_print(f"- {rule['name']}")
        else:
            self.log_print(f"{CRLF}No rules only in {rule_set_one_name}")

        if differences['rules_only_in_2']:
            self.log_print(f"{CRLF}Rules only in {rule_set_two_name} set:")
            for rule in differences['rules_only_in_2']:
                self.log_print(f"- {rule['name']}")
        else:
            self.log_print(f"{CRLF}No rules only in {rule_set_two_name} set")

        if differences['modified_rules']:
            self.log_print(f"{CRLF}Modified rules:")
            for name, rules in differences['modified_rules'].items():
                self.log_print(f"- {name} has differences")
                # Print the differences between the two rules
                self.log_print(f"  {rule_set_one_name} rule: {json.dumps(rules['rules1'], indent=2)}", DEBUG)
                self.log_print(f"  {rule_set_two_name} rule: {json.dumps(rules['rules2'], indent=2)}", DEBUG)
            return False
        else:
            self.log_print(f"{CRLF}No modified rules found", DEBUG)
            return True

    # NOTE: tried to get the outlook junk email options and lists, but could not get it to work
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

    def get_outlook_rules(self):    # no longer in use - YAML rules file is used
        """
        Convert Outlook rules to JSON format with comprehensive error checking.
        Returns a list of rule dictionaries with all available properties.
        """
        rules_json = []
        rules_dict = {}
        timestamp = datetime.now().isoformat()

        try:
            # Get all rules that start with the subset name
            self.log_print("Importing Outlook rules and converting to JSON format...")
            outlook_rules_raw = self.outlook.Session.DefaultStore.GetRules()
            if outlook_rules_raw is None:
                self.log_print("Error: No rules found in Outlook. Ensure rules are configured.")
                return []
            outlook_rules = [rule for rule in outlook_rules_raw if rule.Name.startswith(OUTLOOK_RULES_SUBSET)]
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

            return json.loads(json.dumps(rules_json, indent=2, default=str))

        except Exception as e:
            self.log_print(f"Error accessing Outlook rules: {str(e)}")
            return json.dumps({"error": str(e)})

    def get_safe_senders_rules(self, rules_file=YAML_RULES_SAFE_SENDERS_FILE):
        """
        Read safe senders from YAML file and return as JSON object.
        The safe_senders YAML file contains a list of patterns that can be email addresses or domains.

        Args:
            rules_json: Not used, kept for backward compatibility

        Returns:
            list: List of safe sender patterns, or empty list if file not found/error
        """

        self.log_print("Importing safe senders from YAML file...")
        safe_senders = []
        result = []

        try:
            if not os.path.exists(rules_file):
                self.log_print(f"Safe senders YAML file not found: {rules_file}")
                return safe_senders

            # Read YAML file and convert to Python JSON object per rules_safe_senders.proto definition
            # The YAML file should contain a list of safe senders or a dictionary with a "safe_senders" key
            # where safe_senders[safe_senders] is a list of strings that hold regex pattern strings
            with open(YAML_RULES_SAFE_SENDERS_FILE, 'r', encoding='utf-8') as yaml_file:
                safe_senders = yaml.safe_load(yaml_file)

            if not safe_senders:    # check if file was empty or did not load correctly
                self.log_print("No content found in YAML file")
                return result

            # Extract safe senders list from YAML structure per rules_safe_senders.proto definition
            if isinstance(safe_senders, dict) and 'safe_senders' in safe_senders:
                self.log_print(f"Successfully imported {len(safe_senders['safe_senders'])} safe senders from YAML file")
                self.log_print(f"Safe senders (first 5): {safe_senders['safe_senders'][:5]}")
            elif isinstance(safe_senders, list):
                self.log_print(f"ERROR:  Safe_senders imported as a list from YAML file - need to resolve")
            else:
                self.log_print("No 'safe_senders' key found in YAML file")

            result = json.loads(json.dumps(safe_senders, default=str))

            return result

        except Exception as e:
            self.log_print(f"Error importing safe senders from YAML: {str(e)}")
            self.log_print(f"Error details: {str(e.__class__.__name__)}")
            self.log_print(f"Traceback: {traceback.format_exc()}")
            return result

    def get_yaml_rules(self, rules_file=YAML_RULES_FILE):
        """Import rules from yaml file and return as JSON object (not string)"""
        #*** UPdate to use .proto file
        self.log_print("Importing rules from YAML file...")
        try:
            if not os.path.exists(rules_file):
                self.log_print(f"Rules YAML file not found: {rules_file}")
                return []

            # Read YAML file and convert to Python object
            with open(rules_file, 'r', encoding='utf-8') as yaml_file:
                yaml_content = yaml.safe_load(yaml_file)

            if not yaml_content:
                self.log_print("No rules found in YAML file")
                return []

            # Handle new format with version and settings
            if isinstance(yaml_content, dict) and "rules" in yaml_content:
                # Extract the rules array from the new format
                rules = yaml_content["rules"]
                # Preserve other top-level elements like version and settings
                result = yaml_content
            else:
                # Handle old format where rules were at the top level
                rules = yaml_content if isinstance(yaml_content, list) else [yaml_content]
                result = {"rules": rules}

            self.log_print(f"Successfully imported {len(rules)} rules from YAML file")

            # Convert to JSON using json.dumps and json.loads to ensure consistent structure
            result_json = json.loads(json.dumps(result, default=str))
            return result_json

        except Exception as e:
            self.log_print(f"Error importing rules from YAML: {str(e)}")
            self.log_print(f"Error details: {str(e.__class__.__name__)}")
            import traceback
            self.log_print(f"Traceback: {traceback.format_exc()}")
            return []

    def export_safe_senders_to_yaml(self, rules_json=None, rules_file=YAML_RULES_SAFE_SENDERS_FILE):
        """Export (updated) safe_senders JSON to yaml file"""
        # Update timestamp for each rule - may not be used
        timestamp_rule = datetime.now().isoformat()

        try:
            if rules_json is None:   #this should never happen
                self.log_print("safe_senders JSON is Empty, do not overwrite safe_senders yaml and exit with error")
                return None

            # Convert rules_json to dict if a string format
            if isinstance(rules_json, str):  # convert to JSON
                rules = json.loads(rules_json)
                self.log_print(f"export_safe_senders: Found safe_senders as string and converted to JSON object")
            elif isinstance(rules_json, dict):
                rules = json.loads(json.dumps(rules_json))
                self.log_print(f"export_safe_senders: Found rsafe_senders JSON is a dict and converted to JSON object")
            else:
                # Ensure consistent structure by using json conversion
                rules = json.loads(json.dumps(rules_json, default=str))
                self.log_print(f"export_safe_senders: Standardized rules JSON structure")

            standardized_rules = rules
            # set to all lowercase and remove whitespace
            standardized_rules["safe_senders"] = [item.lower().strip() for item in standardized_rules["safe_senders"]]
            # remove duplicates from the standardized_rules
            standardized_rules["safe_senders"] = list(dict.fromkeys(standardized_rules["safe_senders"]))
            # Sort the safe_senders list
            standardized_rules["safe_senders"] = sorted(standardized_rules["safe_senders"])

            self.log_print(f"Processing {len(standardized_rules["safe_senders"])} safe_senders rules")
            #self.log_print(f"Show list of standardized_rules safe_senders: {standardized_rules}")

            # 03/31/2025 Harold Kimmey Write json_rules to YAML file
            # Ensure directory exists
            os.makedirs(os.path.dirname(rules_file), exist_ok=True)

            # The below section does the following:
            #   Ensure the JSON rules_json object is a valid JSON object before writing to YAML file
            #   Make a backup of the current yaml_file using format "yaml_backup_yyyymmdd_hhmmss.yaml"
            #   Write out the YAML file using YAML_RULES_FILE_temp.yaml
            #   do a get_yaml_rules() to read the file back in and compare it to the original rules_json object to ensure it is a valid match
            #       can use the compare_rules() function to do this
            #   If the comparison is successful, write out the JSON object to YAML_RULES_FILE (use code below)
            #       Convert JSON object to YAML and write to file
            #   If no errors writing the YAML_RULES_FILE, delete the temp file

            # Create a backup of the current YAML file if it exists
            #*** Change copy to a "backup" directory
            if os.path.exists(rules_file):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = f"{os.path.splitext(rules_file)[0]}_backup_{timestamp}.yaml"
                try:
                    import shutil
                    shutil.copy2(rules_file, backup_file)
                    self.log_print(f"Created backup of existing safe_sneders YAML file: {backup_file}")
                except Exception as e:
                    self.log_print(f"Warning: Could not create backup safe_senders file: {str(e)}")

            try:
                with open(rules_file, 'w', encoding='utf-8') as yaml_file:
                    yaml.dump(standardized_rules, yaml_file, sort_keys=False, default_flow_style=False, default_style='"')
                self.log_print(f"Successfully exported {len(standardized_rules["safe_senders"])} safe_senders to YAML file: {rules_file}")

                # # Clean up - delete temporary file
                # try:
                #     os.remove(temp_file)
                #     self.log_print(f"Deleted temporary safe_senders file: {temp_file}")
                # except Exception as e:
                #     self.log_print(f"Warning: Could not delete temporary safe_senders file: {str(e)}")

                return True

            except Exception as e:
                self.log_print(f"Error writing to temporary safe_senders file: {str(e)}")
                return False

        except Exception as e:
            self.log_print(f"Error exporting safe_senders: {str(e)}")
            self.log_print(f"Error details: {str(e.__class__.__name__)}")
            import traceback
            self.log_print(f"Traceback: {traceback.format_exc()}")
            return False


    def export_rules_to_yaml(self, rules_json=None, rules_file=YAML_RULES_FILE):
        """Export JSON/YAML rules to yaml file"""
        # Update timestamp for each rule
        timestamp = datetime.now().isoformat()

        try:
            if rules_json is None:   #this should never happen
                self.log_print("Rules JSON is Empty, do not overwrite rules_file yaml and exit with error")
                return None

            # Convert rules_json to JSON object if it's a string or dict
            if isinstance(rules_json, str):
                rules = json.loads(rules_json)
                self.log_print(f"export_rules: Found rules_json is a string and converted to JSON object")
            elif isinstance(rules_json, dict):
                rules = json.loads(json.dumps(rules_json))
                self.log_print(f"export_rules: Found rules_json is a dict and converted to JSON object")
            else:
                # Ensure consistent structure by using json conversion
                rules = json.loads(json.dumps(rules_json, default=str))
                self.log_print(f"export_rules: Standardized rules JSON structure")


            for rule in rules:
                if isinstance(rule, dict):
                    # Update last_modified in metadata if present, else create metadata
                    if "metadata" in rule and isinstance(rule["metadata"], dict):
                        rule["metadata"]["last_modified"] = timestamp
                    else:
                        rule["metadata"] = {"last_modified": timestamp} #*** need to change so that this happens when rules change

            # Ensure all string values are properly formatted for YAML export
            def ensure_string_values(obj):
                if isinstance(obj, dict):
                    return {k: ensure_string_values(v) for k, v in obj.items()}
                elif isinstance(obj, list):
                    return [ensure_string_values(item) for item in obj]
                elif obj is None:
                    return ""  # Convert None to empty string
                else:
                    return str(obj)  # Ensure all values are strings

            # Apply string formatting to all rules

            standardized_rules = rules
            standardized_rules["rules"] = ensure_string_values(standardized_rules["rules"])
            # Sort all the condition_types for ease in finding them visually
            for rule in standardized_rules.get("rules", []):
                if "conditions" in rule:
                    for condition_type in ["header", "body", "from", "subject"]:
                        if condition_type in rule["conditions"] and isinstance(rule["conditions"][condition_type], list):
                            # ensure all values in rule["conditions"][condition_type] are lowercase and strip whitespace
                            rule["conditions"][condition_type] = [item.lower().strip() for item in rule["conditions"][condition_type]]
                            rule["conditions"][condition_type] = [ensure_string_values(item) for item in rule["conditions"][condition_type]]
                            # remove duplicates from the condition_type list
                            rule["conditions"][condition_type] = list(dict.fromkeys(rule["conditions"][condition_type]))
                            # Sort the condition_type list
                            rule["conditions"][condition_type] = sorted(rule["conditions"][condition_type])


        #    # remove duplicates from the standardized_rules
        #     standardized_rules["safe_senders"] = list(dict.fromkeys(standardized_rules["safe_senders"]))
        #     # Sort the safe_senders list
        #     standardized_rules["safe_senders"] = sorted(standardized_rules["safe_senders"])

            formatted_output = standardized_rules
            self.log_print(f"Number of rules: {len(rules["rules"])}")
            # self.log_print(f"Show list of rules: {rules["rules"]}")
            self.log_print(f"Number of standardized rules: {len(standardized_rules["rules"])}")
            #self.log_print(f"Show list of standardized_rules: {standardized_rules["rules"]}")
            self.log_print(f"Number of formatted_output: {len(formatted_output["rules"])}")
            #self.log_print(f"Show list of formatted_output: {formatted_output["rules"]}")

            # 03/31/2025 Harold Kimmey Write json_rules to YAML file
            # Ensure directory exists
            os.makedirs(os.path.dirname(rules_file), exist_ok=True)

            # The below section does the following:
            #   Ensure the JSON rules_json object is a valid JSON object before writing to YAML file
            #   Make a backup of the current yaml_file using format "yaml_backup_yyyymmdd_hhmmss.yaml"
            #   Write out the YAML file using YAML_RULES_FILE_temp.yaml
            #   do a get_yaml_rules() to read the file back in and compare it to the original rules_json object to ensure it is a valid match
            #       can use the compare_rules() function to do this
            #   If the comparison is successful, write out the JSON object to YAML_RULES_FILE (use code below)
            #       Convert JSON object to YAML and write to file
            #   If no errors writing the YAML_RULES_FILE, delete the temp file

            # Create a backup of the current YAML file if it exists
            if os.path.exists(rules_file):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = f"{os.path.splitext(rules_file)[0]}_backup_{timestamp}.yaml"
                try:
                    import shutil
                    shutil.copy2(rules_file, backup_file)
                    self.log_print(f"Created backup of existing YAML file: {backup_file}")
                except Exception as e:
                    self.log_print(f"Warning: Could not create backup file: {str(e)}")

            #HK 05/13/25 - Removed use of temporary file and comparison before writing to file
            # Create temporary file path
            # temp_file = f"{os.path.splitext(rules_file)[0]}_temp.yaml"


            # Write to file
            try:

                with open(rules_file, 'w', encoding='utf-8') as yaml_file:
                    # If somehow not a list, write as-is (fallback)
                    #*** try new below to have all strings written with double quotes
                    # temporarily remove #yaml.dump(formatted_output, yaml_file, sort_keys=False, default_flow_style=False)
                    yaml.dump(formatted_output, yaml_file, sort_keys=False, default_flow_style=False, default_style='"', width=4096)
                    self.log_print(f"Successfully exported {len(standardized_rules["rules"])} rules to YAML file: {rules_file}")

                # Clean up - delete temporary file
                # try:
                #     os.remove(temp_file)
                #     self.log_print(f"Deleted temporary file: {temp_file}")
                # except Exception as e:
                #     self.log_print(f"Warning: Could not delete temporary file: {str(e)}")

                return True

            except Exception as e:
                self.log_print(f"Error writing to temporary file: {str(e)}")
                return False

        except Exception as e:
            self.log_print(f"Error exporting rules: {str(e)}")
            self.log_print(f"Error details: {str(e.__class__.__name__)}")
            import traceback
            self.log_print(f"Traceback: {traceback.format_exc()}")
            return False


    def get_rules(self):

        """Get rules from YAML file if available, otherwise from Outlook"""
        # 03/31/2025 Harold Kimmey Changing import rules from CSV to YAML file (easy import/export via JSON/YAML)

        YAML_rules = []
        YAML_rules = self.get_yaml_rules(YAML_RULES_FILE)
        self.log_print(f"Import rules from YAML ({YAML_RULES_FILE})")

        #Stop getting Outlook Rules
        # outlook_rules = []
        # self.log_print(f"Import rules from Outlook")
        # outlook_rules = self.get_outlook_rules()

        # debugging - compare YAML_rules to Outlook_rules and print the differences between them
        # self.output_rules_differences(outlook_rules, "Outlook", YAML_rules, "YAML")

        # debugging - for future runs, no longer use Outlook rules
        #rules = outlook_rules

        safe_senders = self.get_safe_senders_rules(YAML_RULES_SAFE_SENDERS_FILE)

        # **NOT needed - after verifying, these lines can be removed
        # # Extract rules array from dictionary if needed
        # if isinstance(rules, dict) and "rules" in rules:
        #     self.log_print(f"Extracted rules array from dictionary wrapper")
        #     return rules["rules"]
        self.log_print(f"Number of rules: {len(YAML_rules["rules"])}")
        #self.log_print(f"Show list of rules: {YAML_rules["rules"]}")
        self.log_print(f"Number of safe_senders rules: {len(safe_senders["safe_senders"])}")
        #self.log_print(f"Show list of safe_senders rules: {safe_senders["safe_senders"]}")

        # Otherwise, return the rules directly
        return YAML_rules, safe_senders

    def print_rules_summary(self, rules):   # rules should be a JSON object
        """Print a summary of all rules in the yaml file"""
        try:
            # add a check to convert to a JSON object (if it a string or dict)
            if isinstance(rules, str) or isinstance(rules, dict):
                rules = json.loads(json.dumps(rules))

            self.log_print(f"{CRLF}Rules Summary:")
            for rule in rules:
                self.log_print(f"\nRule: {rule['name']} (Enabled: {rule['enabled']})")
                for cond_type, values in rule['conditions'].items():
                    if not isinstance(values, list):
                        values = [values]
                    self.log_print(f"  {cond_type} conditions:")
                    for value in values:
                        self.log_print(f"    - {value}")
                self.log_print(f"  Actions:")
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
        line_with_from = ""  # Initialize with an empty string
        blank = ""

        # Handle case where email_header could be a list
        if isinstance(email_header, list):
            email_header = "\n".join(email_header)

        # Iterate over each element in email_header
        for line in email_header.splitlines():
            if line.lower().startswith("from:"):
                line_with_from = line
                break  # find the first line that starts with "from:" then exit loop

        #   print(f"line_with_from: {line_with_from}")  # Debugging output

        if line_with_from:
            from_domain = re.search(r'@[\w.-]+', line_with_from)
            if from_domain:
                from_domain_str = from_domain.group(0)
                #   print(f"from_domain_str: {from_domain_str}")  # Debugging output
                return from_domain_str

        return blank

    def from_report(self, emails_to_process, emails_added_info, rules_json):
        """
        Generate a report of emails with phishing indicators or no rule matches, including the From domain.

        Args:
            emails_to_process (list): List of emails to process.
            emails_added_info (list): List of dictionaries containing additional information about each email.
        """

        processed_count = 0

        # Print a list for Phishing OR Match=false with From: "@<domain>.<>" so they can be easily added to the rules

        for email in emails_to_process:
            processed_count += 1
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
                    self.log_print(f"{output_string}", level="INFO")
                    simple_print(f"{output_string}")

            except Exception as e:
                self.log_print(f"Error processing match = false email: {str(e)}")

            if (DEBUG) and (processed_count >= DEBUG_EMAILS_TO_PROCESS):
                break  # Stop processing more emails in debug mode, then write the report and prompt for rule updates


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

        processed_count = 0

        # Print a list for Phishing OR Match=false, report body unique URL stubs "/<domain>.<>" and ".<domain>.<>" so they can be easily added to the rules
        #     collect them all first, then determine uniqueness, then print one per line

        for email in emails_to_process:
            processed_count += 1
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

            if (DEBUG) and (processed_count >= DEBUG_EMAILS_TO_PROCESS):
                break  # Stop processing more emails in debug mode, then write the report and prompt for rule updates

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

    def get_safe_input(self, prompt_text, valid_responses=None, isregex=False):
        """
        Get user input with security validation.

        Args:
            prompt_text (str): The text to display to the user
            valid_responses (list, optional): List of valid responses. If None, any non-empty input is valid.
            isregex (bool, optional): Whether to allow regex patterns in the input. Defaults to False.

        Returns:
            str: The validated user input
        """
        while True:
            # Get user input
            user_input = input(prompt_text).strip().lower()

            # Define dangerous patterns based on whether regex is allowed
            if isregex:
                # For regex input, we need to be more permissive but still prevent command injection
                dangerous_patterns = [
                    ';', '--', '/*', '*/', 'union', 'select', 'insert', 'update', 'delete',
                    'drop', 'exec', 'execute', '<script', 'javascript:', 'onerror', 'onload',
                    '$(', '${', '`', '&&', '||', '|'  # Allow < and > for regex character classes
                ]

                # Check if the regex pattern is valid by trying to compile it
                try:
                    import re
                    re.compile(user_input)
                except re.error:
                    print("Invalid regex pattern. Please try again.")
                    continue

            else: #*** if not regex, check for valid responses
                # Standard dangerous patterns for non-regex input
                dangerous_patterns = [
                    ';', '--', '/*', '*/', 'union', 'select', 'insert', 'update', 'delete',
                    'drop', 'exec', 'execute', '<script', 'javascript:', 'onerror', 'onload',
                    '$(', '${', '`', '&&', '||', '|', '>', '<', '&lt;', '&gt;'
                ]

                has_dangerous_pattern = any(pattern in user_input.lower() for pattern in dangerous_patterns)

                if has_dangerous_pattern:
                    print("Invalid input. Please try again.")
                    continue

                # Check if response is valid (only for non-regex input)
                if valid_responses and not isregex and user_input not in valid_responses:
                    print(f"Please enter one of: {', '.join(valid_responses)}")
                    continue

            # For regex mode and valid_responses, we could check if the pattern matches any of the valid responses
            # but typically regex mode would be used without valid_responses constraints

            return user_input


    def prompt_update_rules(self, emails_to_process, emails_added_info, rules_json, safe_senders):
        """
        Prompt user to update rules based on unfiltered emails.

        Args:
            emails_to_process (list): List of emails processed.
            emails_added_info (list): Additional info about processed emails.
            rules_json (list): Current rules in JSON format that may be updated.

        Returns:
            list: Updated rules in JSON format.
        """
        self.log_print(f"{CRLF}Checking for emails that can be added to rules...")
        unfiltered_emails = []

        self.log_print(f"Number of emails to process: {len(emails_to_process)}")

        # Find unfiltered emails (those with match=False)
        for i, email_info in enumerate(emails_added_info):
            if email_info["processed"] and email_info["match"] == False and i < len(emails_to_process):
                unfiltered_emails.append((emails_to_process[i], email_info))

        if not unfiltered_emails:
            self.log_print("No unfiltered emails found to update rules.")
            return rules_json, safe_senders
        else:
            self.log_print(f"Found {len(unfiltered_emails)} unfiltered emails to process for rule updates.")

        self.log_print(f"Found {len(unfiltered_emails)} unfiltered emails. Processing for possible rule updates...")
        simple_print(f"\nBeginning interactive rule update for {len(unfiltered_emails)} unfiltered emails")

        #*** What is the best way to ONLY ask once for each unique from_domain in unfiltered_emails
        # create a list of those added and skip others

        # Process each unfiltered email
        # NOTE:  assumes user will only want to update 1 rule per email
        count = 0

        for email, email_info in unfiltered_emails:
            try:
                rule_updated = False
                count += 1
                #   self.log_print(f"before assigning email_header")
                email_header = email_info["email_header"]
                #   self.log_print(f"for loop email_header: {email_header}")  # Debugging output
                subject = self._sanitize_string(email.Subject)
                self.log_print(f"Subject: {subject}")
                from_email = self._sanitize_string(email.SenderEmailAddress).lower()
                self.log_print(f"From: {from_email}")
                from_domain = self.header_from(email_header)
                self.log_print(f"Domain: {from_domain}")
                unique_urls = self.get_unique_URL_stubs(email.Body) # Extract URLs
                self.log_print(f"Unique URLs: {unique_urls}")

                # if the from_email matches a the safe_senders list, skip this email
                if from_domain in safe_senders["safe_senders"]: # or from_email in safe_senders["safe_senders"]:
                    count += 1
                    self.log_print(f"Skipping email from safe sender: {from_email}")
                    simple_print(f"Skipping email from safe sender: {from_email}")
                    continue

                # if the from_domain matches a rule in rules_json header condition, then print a message and skip this email
                if any(from_domain in rule.get("conditions", {}).get("header", []) for rule in rules_json["rules"]):
                    count += 1
                    self.log_print(f"Skipping email from domain as '{from_domain}' already exists in rules.")
                    simple_print(f"Skipping email from domain as '{from_domain}' already exists in rules.")
                    continue

                # if the from_email matches a rule in rules_json header condition, then print a message and skip this email
                if any(from_email in rule.get("conditions", {}).get("header", []) for rule in rules_json["rules"]):
                    count += 1
                    self.log_print(f"Skipping email from email as '{from_email}' already exists in rules.")
                    simple_print(f"Skipping email from email as '{from_email}' already exists in rules.")
                    continue

                # For now assume user wants to process all non-deleted emails
                #
                # if count > 0:   # Ask if user wants to continue with next email (but not first email)
                #     continue_response = input("\nContinue to next email? (y/n): ").lower()
                #     if continue_response != 'y':
                #         simple_print("Rule update process exited by user")
                #         break

                # Display email details
                print(f"{CRLF}" + "=" * 60)
                print(f"Subject: {subject}")
                print(f"From: {from_email}")
                print(f"Domain: {from_domain}")
                print(f"Unique URLs: {unique_urls}")

                # if unique_urls:
                #     simple_print("\nUnique URLs in body:")
                #     for url in unique_urls[:5]:  # Show just the first 5 to avoid overwhelming output
                #         simple_print(f"  {url}")
                #     if len(unique_urls) > 5:
                #         simple_print(f"  ... and {len(unique_urls) - 5} more")

                response = ""

                # Step 1: Suggest header rule
                #*** need to update to highly secure vetting of input from user
                #*** for the following domains that host individual email addresses, only suggest adding full email address to header rules:
                #   gmail.com, yahoo.com, hotmail.com, outlook.com, aol.com, protonmail.com,
                domains_with_individual_emails = from_domain in [
                    "@gmail.com", "@yahoo.com", "@hotmail.com", "@outlook.com", "@aol.com", "@protonmail.com",
                ]


                if from_domain:
                    if domains_with_individual_emails:
                        # For individual email domains, suggest adding full email address
                        expected_responses = ['e', 's']
                        prompt = f"{CRLF}Add '{from_email}' to SpamAutoDeleteHeader rule or safe_senders? ({'/'.join(expected_responses)}): "
                        response = self.get_safe_input(prompt, expected_responses)
                        from_domain = from_email  # Use full email address for individual domains
                    else:
                        expected_responses = ['d', 's']
                        prompt = f"{CRLF}Add '{from_domain}' or email domain to SpamAutoDeleteHeader rule or safe_senders? ({'/'.join(expected_responses)}): "
                        response = self.get_safe_input(prompt, expected_responses)

                    if response == 'd':
                        # Find the SpamAutoDeleteHeader rule in the rules list and append to its header conditions
                        for rule in rules_json["rules"]:
                            if rule["name"] == "SpamAutoDeleteHeader":
                                if "header" not in rule["conditions"]:
                                    rule["conditions"]["header"] = []
                                rule["conditions"]["header"].append(from_domain)
                                rule_updated = True
                                self.log_print(f"Added '{from_domain}' to SpamAutoDeleteHeader rule")
                                simple_print(f"Added '{from_domain}' to SpamAutoDeleteHeader rule")
                    elif response == 'e':
                        # Add from_email to safe_senders list
                        for rule in rules_json["rules"]:
                            if rule["name"] == "SpamAutoDeleteHeader":
                                if "header" not in rule["conditions"]:
                                    rule["conditions"]["header"] = []
                                rule["conditions"]["header"].append(from_domain)
                                rule_updated = True
                                self.log_print(f"Added '{from_domain}' to SpamAutoDeleteHeader rule")
                                simple_print(f"Added '{from_domain}' to SpamAutoDeleteHeader rule")
                    elif response == 's':
                        # Add from_domain to safe_senders list
                        safe_senders["safe_senders"].append(from_domain)  # working HK 05/18/25
                        self.log_print(f"Added '{from_domain}' to safe_senders list")
                        simple_print(f"Added '{from_domain}' to safe_senders list")
                        rule_updated = True

#**** Confirmed working up to this point 05/19/25 HK
#fix these one at a time and add them back in - for now, skip the other options so we can test out the header rule
                # # Step 2: If user declined header rule, suggest body rule for URLs
                # if subject and rule_updated == False:
                #     response = input(f"{CRLF}Add subject pattern '{subject}' to SpamAutoDeleteSubject rule? (y/n): ").lower()

                #     if response == 'y':
                #         subject_pattern = input(f"{CRLF}Enter pattern:")
                #         # Find the SpamAutoDeleteSubject rule
                #         rule_updated = True
                #         for rule in rules_json:
                #             if rule.get("name") == "SpamAutoDeleteSubject":
                #                 if "subject" not in rule["conditions"]:
                #                     rule["conditions"]["subject"] = []

                #                 # Add subject to subject conditions if not already present
                #                 if subject not in rule["conditions"]["subject"]:
                #                     rule["conditions"]["subject"].append(subject_pattern)
                #                     self.log_print(f"Added '{subject_pattern}' to SpamAutoDeleteSubject rule")
                #                     simple_print(f"Added '{subject_pattern}' to SpamAutoDeleteSubject rule")
                #                 else:
                #                     simple_print(f"'{subject_pattern}' already exists in SpamAutoDeleteSubject rule")
                #                 break

                # # Step 3: Suggest subject rule if neither header nor body rules using URL's found
                # if unique_urls and rule_updated == False: #process URL's
                #     simple_print("{CRLF}The following URL patterns can be added to SpamAutoDeleteBody rule:")
                #     for i, url in enumerate(unique_urls):
                #         simple_print(f"{i+1}. {url}")

                #     url_indices = input("Enter URL numbers to add (comma-separated, or 'all'), or press Enter to skip: ")

                #     if url_indices.lower() == 'all':
                #         selected_urls = unique_urls
                #     elif url_indices:
                #         try:
                #             indices = [int(idx.strip()) - 1 for idx in url_indices.split(',')]
                #             selected_urls = [unique_urls[i] for i in indices if 0 <= i < len(unique_urls)]
                #         except ValueError:
                #             simple_print("Invalid input. No URLs added.")
                #             selected_urls = []
                #     else:
                #         selected_urls = []

                #     if selected_urls:
                #         # Add selected URLs to SpamAutoDeleteBody rule
                #         rule_updated = True
                #         for rule in rules_json:
                #             if rule.get("name") == "SpamAutoDeleteBody":
                #                 if "body" not in rule["conditions"]:
                #                     rule["conditions"]["body"] = []

                #                 for url in selected_urls:
                #                     if url not in rule["conditions"]["body"]:
                #                         rule["conditions"]["body"].append(url)
                #                         self.log_print(f"Added '{url}' to SpamAutoDeleteBody rule")
                #                         simple_print(f"Added '{url}' to SpamAutoDeleteBody rule")

                # if rule_updated == False:  # request body content add if no other adds
                #     print(f"{CRLF}Body Content:")
                #     for line in email.Body.splitlines():
                #         self.log_print(f"Body: {line}")
                #     response = input(f"{CRLF}Add body content to SpamAutoDeleteBody rule? (y/n): ").lower()
                #     if response == 'y':
                #         body_content = input(f"{CRLF}Enter body content:")
                #         # Find the SpamAutoDeleteBody rule
                #         rule_updated = True
                #         for rule in rules_json:
                #             if rule.get("name") == "SpamAutoDeleteBody":
                #                 if "body" not in rule["conditions"]:
                #                     rule["conditions"]["body"] = []

                #                 # Add body content to body conditions if not already present
                #                 if body_content not in rule["conditions"]["body"]:
                #                     rule["conditions"]["body"].append(body_content)
                #                     self.log_print(f"Added '{body_content}' to SpamAutoDeleteBody rule")
                #                     simple_print(f"Added '{body_content}' to SpamAutoDeleteBody rule")
                #                 else:
                #                     simple_print(f"'{body_content}' already exists in SpamAutoDeleteBody rule")
                #                 break

            except Exception as e:
                self.log_print(f"Error processing email for rule updates: {str(e)} {email_header}")
                simple_print(f"Error processing email: {str(e)}")

        self.log_print("Rule update process completed")
        simple_print("\nRule update process completed")
        return rules_json, safe_senders

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

    def process_emails(self, rules_json, safe_senders, days_back=DAYS_BACK_DEFAULT):
        """Process emails based on the rules in the rules_json object"""
        self.log_print(f"\n\nStarting email processing")
        self.log_print(f"Target folder: {self.target_folder.Name}", "DEBUG")
        self.log_print(f"Processing emails from last {days_back} days")

        try:
            # Extract rules array if rules_json is a dictionary with a 'rules' key
            if isinstance(rules_json, dict) and "rules" in rules_json:
                rules = rules_json["rules"]
                # safe_senders = rules_json.get("safe_senders", [])
                self.log_print(f"Processing with {len(rules)} rules and {len(safe_senders["safe_senders"])} safe senders")
            else:
                # Handle direct rules array
                rules = rules_json if isinstance(rules_json, list) else [rules_json]
                safe_senders = []

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
            self.log_print(f"before adding fields to emails_added_info", DEBUG)
            emails_added_info = [{
                "match": False,
                "rule": "",
                "matched_keyword": "",
                "indicators": [],
                "email_header": "",
                "processed": False,
            } for email in emails_to_process]
            self.log_print(f"after adding fields to emails_added_info", DEBUG)

            for email in emails_to_process:
                try:
                    processed_count += 1
                    email_index = emails_to_process.index(email)
                    email_deleted = False
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

                        # Check 'subject' keywords
                        if 'subject' in conditions:
                            if any(keyword.lower() in email.Subject.lower() for keyword in conditions['subject']):
                                match = True
                                matched_keyword = next((keyword for keyword in conditions['subject'] if keyword.lower() in email.Subject.lower()), None)
                                self.log_print(f"Matched keyword in subject: {matched_keyword}")
                                self.log_print(f"Subject: {self._sanitize_string(email.Subject)}")

                        # Check 'body' keywords
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
                        # Check 'header' keywords
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

                        if match and safe_senders and "safe_senders" in safe_senders:
                            if any(sender.lower() in email_header.lower() for sender in safe_senders["safe_senders"]):
                                match = False
                                matched_sender = next((sender for sender in safe_senders["safe_senders"] if sender.lower() in email_header.lower()), None)
                                self.log_print(f"Safe sender matched in header: {matched_sender}")
                                self.log_print(f"Email not processed due to safe sender match")

                        if match and 'from' in exceptions:
                            from_addresses = [addr['address'].lower() for addr in exceptions['from']]
                            if any(addr in email.SenderEmailAddress.lower() for addr in from_addresses):
                                match = False
                                matched_keyword = next((addr['address'] for addr in exceptions['from'] if addr['address'].lower() in email.SenderEmailAddress.lower()), None)
                                self.log_print(f"Exception matched keyword in from address: {matched_keyword}")
                                self.log_print(f"From: {self._sanitize_string(email.SenderEmailAddress)}")

                        # Check subject keywords in exceptions
                        if match and 'subject' in exceptions:
                            if any(keyword.lower() in email.Subject.lower() for keyword in exceptions['subject']):
                                match = False
                                matched_keyword = next((keyword for keyword in exceptions['subject'] if keyword.lower() in email.Subject.lower()), None)
                                self.log_print(f"Exception matched keyword in subject: {matched_keyword}")
                                self.log_print(f"Subject: {self._sanitize_string(email.Subject)}")

                        # Check body keywords in exceptions
                        if match and 'body' in exceptions:
                            if any(keyword.lower() in email.Body.lower() for keyword in exceptions['body']):
                                match = False
                                matched_keyword = next((keyword for keyword in exceptions['body'] if keyword.lower() in email.Body.lower()), None)
                                self.log_print(f"Exception matched keyword in body: {matched_keyword}")
                                self.log_print(f"Body: {self._sanitize_string(email.Body)}")

                        # Check header keywords in exceptions
                        if match and 'header' in exceptions:
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
                            emails_added_info[email_index]["processed"] = True
                        else:
                            emails_added_info[email_index]["match"] = match
                            emails_added_info[email_index]["rule"] = None
                            emails_added_info[email_index]["matched_keyword"] = ""
                            emails_added_info[email_index]["email_header"] = email_header
                            emails_added_info[email_index]["processed"] = True

                        if match:
                            self.log_print(f"Email matches rule: {rule['name']}")
                            # Perform actions based on the rule
                            actions = rule['actions']
                            self.log_print(f"Performing actions: {actions}")

                            if 'assign_to_category' in actions and actions['assign_to_category']['category_name']:
                                try: # to assign category based on rule name
                                    category_name = actions['assign_to_category']['category_name']
                                    self.assign_category_to_email_with_retry(email, category_name)
                                    self.log_print(f"Email assigned to category '{category_name}'", "DEBUG")
                                except Exception as e:
                                    self.log_print(f"Error assigning category to email: {str(e)}")
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

                                try: # to delete email
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

                    if (DEBUG) and (processed_count >= DEBUG_EMAILS_TO_PROCESS):
                        self.log_print(f"Debug mode: Stopping after {DEBUG_EMAILS_TO_PROCESS} emails")
                        break  # Stop processing more emails in debug mode, then write the report and prompt for rule updates

                except Exception as e:
                    self.log_print(f"Error processing email: {str(e)}")


            # Print a list for Phishing OR Match=false, report body unique URL stubs "/<domain>.<>" and ".<domain>.<>" so they can be easily added to the rules
            #     collect them all first, then determine uniqueness, then print one per line
            self.log_print(f"\nProcessing Report of URL's from phishing or match = False")
            self.URL_report(emails_to_process, emails_added_info)

            # Print a list for Phishing OR Match=false with From: "@<domain>.<>" so they can be easily added to the rules
            self.log_print(f"\nProcessing Report of From's from phishing or match = False")
            self.from_report(emails_to_process, emails_added_info, rules_json)

            # After processing all emails, prompt for rule updates based on unfiltered emails
            if processed_count > 0:
                self.log_print(f"{CRLF}Prompting for rule updates based on unfiltered emails...")
                rules_json, safe_senders = self.prompt_update_rules(emails_to_process, emails_added_info, rules_json, safe_senders)

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
    agent = OutlookSecurityAgent()  # setup for calling functions in class OutlookSecurityAgent

    try:
        simple_print(f"\n=============================================================")
        simple_print(f"Starting Outlook Security Agent at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}")
        simple_print(f"This will make changes")
        simple_print(f"Check the {OUTLOOK_SECURITY_LOG} for detailed information")
        agent.log_print(f"\n=============================================================")
        agent.log_print(f"Starting Outlook Security Agent at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}")
        agent.log_print(f"This will make changes")
        agent.log_print(f"Check the {OUTLOOK_SECURITY_LOG} for detailed information")

        rules_json, safe_senders = agent.get_rules()

        # Process last N days of emails - see DAYS_BACK_DEFAULT
        agent.log_print(f"{CRLF}Begin email analysis{CRLF}")

        agent.process_emails(rules_json, safe_senders)

        agent.log_print(f"{CRLF}End email analysis{CRLF}")

        # Export rules every time (saving copies to backups to Archive directory)
        agent.export_rules_to_yaml(rules_json)

        #*** run export_safe_senders_to_yaml function to export safe_senders to YAML
        #***   ensure they are written in the format they of the current file - see merge_safe_senders.py
        #***   create backup copies, like rules_json
        #***   after running verify they can be read
        agent.export_safe_senders_to_yaml(safe_senders)

        simple_print(f"Execution complete at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}. Check the log file for detailed analysis:\n{OUTLOOK_SECURITY_LOG}")
        simple_print(f"=============================================================\n")
        agent.log_print(f"Execution complete at {datetime.now().strftime('%m/%d/%Y %I:%M:%S %p')}. Check the log file for detailed analysis:\n{OUTLOOK_SECURITY_LOG}")
        agent.log_print(f"============================================================={CRLF}{CRLF}")

    except Exception as e:
        simple_print(f"\nError: {str(e)}")
        logging.error(f"Main execution error: {str(e)}")

if __name__ == "__main__":
    main()
