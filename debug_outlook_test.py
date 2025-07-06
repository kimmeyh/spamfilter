import unittest
import os
import sys

# Add the directory containing withOutlookRulesCSV.py to the Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from withOutlookRulesYAML import OutlookSecurityAgent

class TestDebugOutlookConnection(unittest.TestCase):

    def setUp(self):
        """Set up test environment with detailed debugging"""
        print(f"\nDEBUG setUp: Working directory: {os.getcwd()}")
        print(f"DEBUG setUp: Python path contains: {[p for p in sys.path if 'OutlookMailSpamFilter' in p]}")
        
        try:
            print("DEBUG setUp: Creating OutlookSecurityAgent...")
            self.agent = OutlookSecurityAgent(debug_mode=True)
            print(f"DEBUG setUp: SUCCESS - Agent created with email: {self.agent.email_address}")
            print(f"DEBUG setUp: SUCCESS - Found {len(self.agent.target_folders)} target folders")
            for i, folder in enumerate(self.agent.target_folders):
                print(f"DEBUG setUp: Folder {i+1}: {folder.Name}")
            
            print("DEBUG setUp: Getting rules...")
            self.rules_json, self.safe_senders = self.agent.get_rules()
            print(f"DEBUG setUp: SUCCESS - Rules retrieved: {type(self.rules_json)}")
            print(f"DEBUG setUp: SUCCESS - Safe senders retrieved: {type(self.safe_senders)}")
            
            self.yaml_file = "test_rules_debug.yaml"
            print("DEBUG setUp: Setup completed successfully")
            
        except ValueError as e:
            print(f"DEBUG setUp: ValueError caught: {e}")
            if "Could not find any of the specified folders" in str(e):
                print("DEBUG setUp: This is the folder detection error - skipping test")
                self.skipTest(f"Skipping test due to missing Outlook folders: {e}")
            else:
                print("DEBUG setUp: Different ValueError - re-raising")
                raise
        except Exception as e:
            print(f"DEBUG setUp: Other exception: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            raise

    def tearDown(self):
        """Clean up test environment"""
        if hasattr(self, 'yaml_file') and os.path.exists(self.yaml_file):
            os.remove(self.yaml_file)

    def test_export_import_rules_yaml(self):
        """Test exporting and importing rules to/from YAML"""
        print("\nDEBUG test: Starting export/import test")
        
        # Export rules to YAML
        print("DEBUG test: Exporting rules to YAML...")
        self.agent.export_rules_to_yaml(self.rules_json, self.yaml_file)
        print("DEBUG test: Export completed")

        # Import rules from YAML
        print("DEBUG test: Importing rules from YAML...")
        imported_rules = self.agent.get_yaml_rules(self.yaml_file)
        print(f"DEBUG test: Import completed, type: {type(imported_rules)}")

        # Compare the original rules and the imported rules, ignoring timestamp differences
        print("DEBUG test: Comparing rules (ignoring timestamps)...")
        
        def remove_timestamps(data):
            """Recursively remove last_modified timestamps for comparison"""
            if isinstance(data, dict):
                cleaned = {}
                for key, value in data.items():
                    if key == "last_modified":
                        continue  # Skip timestamp fields
                    else:
                        cleaned[key] = remove_timestamps(value)
                return cleaned
            elif isinstance(data, list):
                return [remove_timestamps(item) for item in data]
            else:
                return data
        
        original_clean = remove_timestamps(self.rules_json)
        imported_clean = remove_timestamps(imported_rules)
        
        try:
            self.assertEqual(original_clean, imported_clean, "The imported rules do not match the original rules (ignoring timestamps)")
            print("DEBUG test: Rules match (ignoring timestamps)!")
        except AssertionError as e:
            print(f"DEBUG test: Rules don't match: {e}")
            # If they don't match, print out the differences
            differences = self.agent.compare_rules(original_clean, imported_clean)
            print("Differences found:")
            import json
            print(json.dumps(differences, indent=2))
            raise

if __name__ == '__main__':
    unittest.main(verbosity=2)
