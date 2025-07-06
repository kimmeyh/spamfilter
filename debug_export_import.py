#!/usr/bin/env python3
"""Debug script to test export/import functionality directly"""

import os
import sys
import json

# Add the directory containing withOutlookRulesCSV.py to the Python path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from withOutlookRulesYAML import OutlookSecurityAgent

def debug_export_import():
    print("=== Export/Import Debug Test ===")
    
    try:
        # Create agent
        print("1. Creating agent...")
        agent = OutlookSecurityAgent(debug_mode=True)
        print(f"   ✓ Agent created successfully")
        
        # Get rules
        print("2. Getting rules...")
        rules_data, safe_senders_data = agent.get_rules()
        print(f"   ✓ Rules retrieved: {type(rules_data)}")
        print(f"   ✓ Rules structure: {list(rules_data.keys()) if isinstance(rules_data, dict) else 'Not a dict'}")
        
        if isinstance(rules_data, dict) and 'rules' in rules_data:
            print(f"   ✓ Found {len(rules_data['rules'])} rules")
        else:
            print(f"   ⚠ Unexpected rules structure: {rules_data}")
            return
        
        # Test export
        test_file = "debug_test_export.yaml"
        print(f"3. Exporting to {test_file}...")
        try:
            export_result = agent.export_rules_to_yaml(rules_data, test_file)
            print(f"   Export result: {export_result}")
        except Exception as e:
            print(f"   Export FAILED with exception: {e}")
            import traceback
            traceback.print_exc()
            return
        
        # Check if file exists
        if os.path.exists(test_file):
            print(f"   ✓ File created successfully")
            file_size = os.path.getsize(test_file)
            print(f"   File size: {file_size} bytes")
        else:
            print(f"   ✗ File NOT created")
            return
        
        # Test import
        print(f"4. Importing from {test_file}...")
        imported_data = agent.get_yaml_rules(test_file)
        print(f"   ✓ Import completed: {type(imported_data)}")
        
        if isinstance(imported_data, dict):
            print(f"   ✓ Imported structure: {list(imported_data.keys())}")
            if 'rules' in imported_data:
                print(f"   ✓ Found {len(imported_data['rules'])} imported rules")
            else:
                print(f"   ⚠ No 'rules' key in imported data")
        elif isinstance(imported_data, list):
            print(f"   ⚠ Imported as list with {len(imported_data)} items")
        else:
            print(f"   ⚠ Unexpected import type: {type(imported_data)}")
            print(f"   Content: {imported_data}")
        
        # Compare
        print("5. Comparing data...")
        if rules_data == imported_data:
            print("   ✓ Data matches perfectly!")
        else:
            print("   ✗ Data does NOT match")
            print(f"   Original type: {type(rules_data)}")
            print(f"   Imported type: {type(imported_data)}")
            
            # Show first difference
            if isinstance(rules_data, dict) and isinstance(imported_data, dict):
                for key in rules_data:
                    if key not in imported_data:
                        print(f"   Missing key in imported: {key}")
                    elif rules_data[key] != imported_data[key]:
                        print(f"   Different value for key '{key}'")
                        print(f"     Original: {type(rules_data[key])} - {str(rules_data[key])[:100]}...")
                        print(f"     Imported: {type(imported_data[key])} - {str(imported_data[key])[:100]}...")
                        break
        
        # Cleanup
        if os.path.exists(test_file):
            os.remove(test_file)
            print(f"   Cleaned up {test_file}")
        
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_export_import()
