import unittest
import os
import json
import sys

# Add the directory containing withOutlookRulesCSV.py to the Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from withOutlookRulesYAML import OutlookSecurityAgent

class TestOutlookRulesCSV(unittest.TestCase):

    def setUp(self):
        """Set up test environment"""
        try:
            self.agent = OutlookSecurityAgent()
            self.rules_json = self.agent.get_rules()
            self.yaml_file = "test_rules.yaml"
        except ValueError as e:
            if "Could not find any of the specified folders" in str(e):
                # This is expected in test environment without real Outlook folders
                self.skipTest(f"Skipping test due to missing Outlook folders: {e}")
            else:
                raise

    def tearDown(self):
        """Clean up test environment"""
        if os.path.exists(self.yaml_file):
            os.remove(self.yaml_file)

    def test_export_import_rules_yaml(self):
        """Test exporting and importing rules to/from YAML"""
        # Export rules to YAML
        self.agent.export_rules_to_yaml(self.rules_json, self.yaml_file)

        # Import rules from YAML
        imported_rules = self.agent.import_rules_yaml(self.yaml_file)

        # Compare the original rules and the imported rules
        self.assertEqual(self.rules_json, imported_rules, "The imported rules do not match the original rules")

        # If they don't match, print out the differences
        if self.rules_json != imported_rules:
            differences = self.agent.compare_rules(self.rules_json, imported_rules)
            print("Differences found:")
            print(json.dumps(differences, indent=2))

if __name__ == '__main__':
    unittest.main()
