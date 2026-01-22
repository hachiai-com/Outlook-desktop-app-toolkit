#!/usr/bin/env python3
"""
Helper script to find Outlook email account IDs
Run this to discover the exact account names to use with the toolkit
"""
import win32com.client
import pythoncom
import sys


def find_accounts():
    """Find and display all Outlook email account IDs"""
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Connect to Outlook
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("\n" + "="*60)
        print("Available Outlook Email Accounts")
        print("="*60)
        print("\nUse these EXACT names as 'email_account' parameter:\n")
        
        accounts = []
        for folder in namespace.Folders:
            account_name = folder.Name
            accounts.append(account_name)
            
            # Try to get inbox info
            try:
                inbox = folder.Folders("Inbox")
                item_count = inbox.Items.Count
                print(f"  📧 {account_name}")
                print(f"     Inbox items: {item_count}")
                
                # Try to get more account details
                try:
                    # Get account from namespace
                    for account in namespace.Accounts:
                        if account.DisplayName == account_name or account.SmtpAddress == account_name:
                            print(f"     SMTP: {account.SmtpAddress}")
                            break
                except:
                    pass
                
                print()
            except Exception as e:
                print(f"  📧 {account_name}")
                print(f"     ⚠️  Could not access Inbox: {str(e)}")
                print()
        
        print("="*60)
        print(f"\nFound {len(accounts)} account(s)")
        print("\n💡 Tip: Copy the exact account name (case-sensitive) for use in toolkit calls")
        
        # Cleanup
        pythoncom.CoUninitialize()
        
        return accounts
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        print("\nTroubleshooting:")
        print("1. Ensure Outlook desktop app is running")
        print("2. Ensure Outlook is not in safe mode")
        print("3. Check that at least one email account is configured")
        sys.exit(1)


if __name__ == "__main__":
    find_accounts()
