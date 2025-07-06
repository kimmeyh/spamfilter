"""
Diagnostic test to understand the OutlookSecurityAgent folder detection issue
"""

import pytest
import os
import sys
from pathlib import Path

# Import the main class
from withOutlookRulesYAML import OutlookSecurityAgent, EMAIL_ADDRESS, EMAIL_BULK_FOLDER_NAMES


def test_diagnostic_outlook_connection():
    """Diagnostic test to understand folder detection"""
    print(f"\n=== DIAGNOSTIC TEST ===")
    print(f"Current working directory: {os.getcwd()}")
    print(f"Python executable: {sys.executable}")
    print(f"EMAIL_ADDRESS: {EMAIL_ADDRESS}")
    print(f"EMAIL_BULK_FOLDER_NAMES: {EMAIL_BULK_FOLDER_NAMES}")
    
    try:
        print("\nAttempting to create OutlookSecurityAgent in test mode...")
        agent = OutlookSecurityAgent(debug_mode=True, test_mode=True)
        print(f"✅ SUCCESS: Agent created in test mode")
        print(f"   Email: {agent.email_address}")
        print(f"   Target folders found: {len(agent.target_folders)} (test mode allows 0)")
        for i, folder in enumerate(agent.target_folders):
            print(f"     Folder {i+1}: {folder.Name}")
        
        print("\nAttempting to get rules...")
        rules_data, safe_senders = agent.get_rules()
        print(f"✅ SUCCESS: Rules retrieved")
        print(f"   Rules type: {type(rules_data)}")
        print(f"   Safe senders type: {type(safe_senders)}")
        
        assert agent is not None
        # In test mode, we don't require folders to exist
        print("✅ Test passed: Agent created successfully in test mode")
        
    except ValueError as e:
        print(f"❌ ValueError: {e}")
        # Let's see if we can get more details about the error
        print("\nTrying to diagnose the issue...")
        
        # Try to import win32com
        try:
            import win32com.client
            print("✅ win32com.client import successful")
            
            # Try to get Outlook application
            outlook = win32com.client.Dispatch("Outlook.Application")
            print("✅ Outlook.Application dispatch successful")
            
            # Try to get namespace
            namespace = outlook.GetNamespace("MAPI")
            print("✅ MAPI namespace successful")
            
            # Try to get accounts
            accounts = namespace.Accounts
            print(f"✅ Found {accounts.Count} accounts")
            
            # Look for our specific account
            target_account = None
            for account in accounts:
                print(f"   Account: {account.DisplayName}")
                if account.DisplayName == EMAIL_ADDRESS:
                    target_account = account
                    print(f"   ✅ Found target account: {EMAIL_ADDRESS}")
                    break
            
            if target_account:
                # Try to get delivery store
                delivery_store = target_account.DeliveryStore
                print(f"✅ Got delivery store: {delivery_store.DisplayName}")
                
                # Try to get root folder
                root_folder = delivery_store.GetRootFolder()
                print(f"✅ Got root folder: {root_folder.Name}")
                
                # List all folders
                print("\nAll folders in account:")
                def list_folders(folder, indent=0):
                    print("  " * indent + f"- {folder.Name}")
                    for subfolder in folder.Folders:
                        list_folders(subfolder, indent + 1)
                
                list_folders(root_folder)
                
            else:
                print(f"❌ Could not find account: {EMAIL_ADDRESS}")
                
        except Exception as e2:
            print(f"❌ Error during diagnosis: {e2}")
            import traceback
            traceback.print_exc()
        
        # Re-raise the original error
        raise
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    test_diagnostic_outlook_connection()
