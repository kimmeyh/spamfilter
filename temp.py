import os
import sys
import yaml
import difflib
import json
from datetime import datetime

# Import the OutlookSecurityAgent from withOutlookRulesYAML
from withOutlookRulesYAML import OutlookSecurityAgent, YAML_RULES_FILE

def normalize_yaml(yaml_content):
    """Convert YAML to normalized dictionary to ensure consistent comparison"""
    try:
        if isinstance(yaml_content, str):
            # Parse YAML string to Python objects
            return yaml.safe_load(yaml_content)
        return yaml_content
    except Exception as e:
        print(f"Error normalizing YAML: {e}")
        return None

def compare_yaml_files(file1, file2, detailed=True):
    """
    Compare two YAML files and report differences

    Args:
        file1: Path to first YAML file
        file2: Path to second YAML file
        detailed: If True, show detailed line-by-line differences

    Returns:
        bool: True if files are equivalent, False otherwise
    """
    print(f"\nComparing YAML files:")
    print(f"File 1: {file1}")
    print(f"File 2: {file2}")

    # Check if files exist
    if not os.path.exists(file1):
        print(f"Error: File not found: {file1}")
        return False
    if not os.path.exists(file2):
        print(f"Error: File not found: {file2}")
        return False

    # Read files
    with open(file1, 'r', encoding='utf-8') as f1, open(file2, 'r', encoding='utf-8') as f2:
        content1 = f1.read()
        content2 = f2.read()

    # Parse YAML content to Python objects for structure comparison
    yaml1 = normalize_yaml(content1)
    yaml2 = normalize_yaml(content2)

    # Compare YAML structure (ignoring formatting differences)
    structure_match = json.dumps(yaml1, sort_keys=True) == json.dumps(yaml2, sort_keys=True)

    # Line-by-line comparison for detailed report
    if detailed and not structure_match:
        print("\nDetailed differences:")
        lines1 = content1.splitlines()
        lines2 = content2.splitlines()

        diff = difflib.unified_diff(
            lines1, lines2,
            fromfile=file1, tofile=file2,
            lineterm=''
        )

        diff_lines = list(diff)
        if diff_lines:
            for line in diff_lines:
                print(line)
        else:
            print("(No line-by-line differences detected)")

    if structure_match:
        print("RESULT: YAML files are structurally equivalent (content matches)")
        return True
    else:
        print("RESULT: YAML files are different")
        return False

def test_yaml_rules():
    """Test loading and exporting YAML rules"""
    print("Starting YAML rules test")

    # Create instance of OutlookSecurityAgent
    try:
        agent = OutlookSecurityAgent(debug_mode=True)
        print("Successfully created OutlookSecurityAgent instance")
    except Exception as e:
        print(f"Error creating OutlookSecurityAgent: {e}")
        return False

    try:
        # Step 1: Load rules from YAML file
        print("\nStep 1: Loading rules from YAML file")
        rules_json, safe_senders = agent.get_rules()
        print(f"Loaded {len(rules_json.get('rules', [])) if isinstance(rules_json, dict) else len(rules_json)} rules")
        print(f"Loaded {len(safe_senders)} safe senders")

        # Step 2: Export rules to YAML file
        print("\nStep 2: Exporting rules to YAML file")
        success = agent.export_rules_to_yaml(rules_json)
        if not success:
            print("Failed to export rules to YAML")
            return False
        print("Successfully exported rules to YAML")

        # Step 3: Find the backup file created during export
        print("\nStep 3: Locating backup file")
        backup_files = [f for f in os.listdir(os.path.dirname(YAML_RULES_FILE))
                       if f.startswith(os.path.basename(YAML_RULES_FILE).split('.')[0] + "_backup_")
                       and f.endswith('.yaml')]

        if not backup_files:
            print("No backup files found")
            return False

        # Sort by modification time to get the most recent backup
        backup_files.sort(key=lambda x: os.path.getmtime(os.path.join(os.path.dirname(YAML_RULES_FILE), x)), reverse=True)
        latest_backup = os.path.join(os.path.dirname(YAML_RULES_FILE), backup_files[0])
        print(f"Latest backup file: {latest_backup}")

        # Step 4: Compare the original and exported YAML files
        print("\nStep 4: Comparing original and exported YAML files")
        files_match = compare_yaml_files(latest_backup, YAML_RULES_FILE)

        return files_match

    except Exception as e:
        print(f"Error during test: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print(f"=== YAML Rules Test - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    result = test_yaml_rules()
    print(f"\nTest {'PASSED' if result else 'FAILED'}")
    sys.exit(0 if result else 1)
