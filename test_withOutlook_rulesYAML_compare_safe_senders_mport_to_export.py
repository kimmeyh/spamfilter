import os
import sys
import yaml
import difflib
import json
from datetime import datetime

# Import the OutlookSecurityAgent from withOutlookRulesYAML
from withOutlookRulesYAML import OutlookSecurityAgent, YAML_RULES_SAFE_SENDERS_FILE

def test_safe_senders_yaml():
    """Test loading and exporting safe senders YAML"""
    print("Starting Safe Senders YAML test")

    # Create instance of OutlookSecurityAgent
    try:
        agent = OutlookSecurityAgent(debug_mode=True)
        print("Successfully created OutlookSecurityAgent instance")
    except Exception as e:
        print(f"Error creating OutlookSecurityAgent: {e}")
        return False

    try:
        # Step 1: Load safe senders from YAML file
        print("\nStep 1: Loading safe senders from YAML file")
        rules_json, safe_senders = agent.get_rules()
        print(f"Loaded {len(safe_senders)} safe senders")

        # Step 2: Export safe senders to YAML file using the agent's function
        print("\nStep 2: Exporting safe senders to YAML file")

        # Create a backup of the original file before we modify it
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = f"{os.path.splitext(YAML_RULES_SAFE_SENDERS_FILE)[0]}_test_backup_{timestamp}.yaml"

        try:
            with open(YAML_RULES_SAFE_SENDERS_FILE, 'r', encoding='utf-8') as src, open(backup_file, 'w', encoding='utf-8') as dst:
                dst.write(src.read())
            print(f"Created backup file: {backup_file}")
        except Exception as e:
            print(f"Error creating backup: {e}")
            return False

        # Create test file path for comparison
        test_file = f"{os.path.splitext(YAML_RULES_SAFE_SENDERS_FILE)[0]}_test.yaml"

        # Use the agent's export function
        success = agent.export_safe_senders_to_yaml(safe_senders, test_file)
        if not success:
            print("Failed to export safe senders to YAML")
            return False
        print("Successfully exported safe senders to YAML")

        # Step 3: Compare the original and exported files
        print("\nStep 3: Comparing original and test YAML files")

        # Read original and test files
        with open(YAML_RULES_SAFE_SENDERS_FILE, 'r', encoding='utf-8') as f1, open(test_file, 'r', encoding='utf-8') as f2:
            content1 = f1.read()
            content2 = f2.read()

        # Parse YAML content to Python objects
        yaml1 = yaml.safe_load(content1)
        yaml2 = yaml.safe_load(content2)

        # Compare structures (ignoring formatting)
        if 'safe_senders' in yaml1 and 'safe_senders' in yaml2:
            # Extract just the safe sender lists for comparison
            senders1 = set(yaml1['safe_senders'])
            senders2 = set(yaml2['safe_senders'])

            if senders1 == senders2:
                print("RESULT: Safe sender lists are equivalent")
                result = True
            else:
                print(f"RESULT: Safe sender lists differ")
                print(f"  Only in original: {len(senders1 - senders2)}")
                print(f"  Only in test file: {len(senders2 - senders1)}")
                result = False
        else:
            print("RESULT: YAML structure is different - missing 'safe_senders' key")
            result = False

        # # Clean up test file
        # if os.path.exists(test_file):
        #     os.remove(test_file)
        #     print(f"Removed test file: {test_file}")

        return result

    except Exception as e:
        print(f"Error during test: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print(f"=== Safe Senders YAML Test - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    result = test_safe_senders_yaml()
    print(f"\nTest {'PASSED' if result else 'FAILED'}")
    sys.exit(0 if result else 1)
