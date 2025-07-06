#!/usr/bin/env python3
"""
Debug script to list all available folders in Outlook account
"""

import win32com.client

def list_all_folders():
    """List all folders in all Outlook accounts"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("=== OUTLOOK ACCOUNTS AND FOLDERS ===")
        
        # Loop through all accounts
        for account in outlook.Session.Accounts:
            print(f"\nAccount: {account.SmtpAddress}")
            print(f"Display Name: {account.DisplayName}")
            
            try:
                # Get the root folder for this account
                root_folder = namespace.Folders(account.DeliveryStore.DisplayName)
                print(f"Root Folder: {root_folder.Name}")
                
                # List all folders in this account
                print("Available folders:")
                list_folders_recursive(root_folder, indent="  ")
                
            except Exception as e:
                print(f"  Error accessing folders: {e}")
                
    except Exception as e:
        print(f"Error: {e}")

def list_folders_recursive(folder, indent=""):
    """Recursively list all folders"""
    try:
        for subfolder in folder.Folders:
            print(f"{indent}{subfolder.Name}")
            # Recursively list subfolders
            list_folders_recursive(subfolder, indent + "  ")
    except Exception as e:
        print(f"{indent}Error listing subfolders: {e}")

if __name__ == "__main__":
    list_all_folders()
