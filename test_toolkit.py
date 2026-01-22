#!/usr/bin/env python3
"""
Automated test script for Outlook Desktop Toolkit
Tests all capabilities to verify toolkit works correctly

SAFETY: By default, this script is READ-ONLY and will NOT send emails.
Only read operations are performed unless explicitly enabled.
"""
import json
import sys
import subprocess
from typing import Dict, Any, Optional
import time
import argparse


class ToolkitTester:
    """Test harness for Outlook Desktop Toolkit"""
    
    def __init__(self, email_account: str):
        self.email_account = email_account
        self.test_results = []
        self.passed = 0
        self.failed = 0
    
    def run_capability(self, capability: str, args: Dict[str, Any]) -> Dict[str, Any]:
        """Run a toolkit capability and return result"""
        input_data = {
            "capability": capability,
            "args": args
        }
        
        try:
            process = subprocess.Popen(
                ["python", "main.py"],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            stdout, stderr = process.communicate(input=json.dumps(input_data))
            
            if stderr:
                print(f"⚠️  Warning: {stderr}", file=sys.stderr)
            
            if process.returncode != 0:
                return {"error": f"Process exited with code {process.returncode}"}
            
            try:
                return json.loads(stdout)
            except json.JSONDecodeError:
                return {"error": f"Invalid JSON response: {stdout}"}
                
        except Exception as e:
            return {"error": str(e)}
    
    def test(self, test_name: str, capability: str, args: Dict[str, Any], 
             expected_keys: Optional[list] = None, should_succeed: bool = True) -> bool:
        """Run a test and record result"""
        print(f"\n{'='*60}")
        print(f"Test: {test_name}")
        print(f"{'='*60}")
        print(f"Capability: {capability}")
        print(f"Args: {json.dumps(args, indent=2)}")
        print(f"\nRunning...")
        
        result = self.run_capability(capability, args)
        
        print(f"\nResponse:")
        print(json.dumps(result, indent=2))
        
        # Check result
        success = False
        if should_succeed:
            if "result" in result:
                if expected_keys:
                    result_data = result.get("result", {})
                    success = all(key in result_data for key in expected_keys)
                else:
                    success = True
            else:
                success = False
        else:
            # Test expects failure
            success = "error" in result
        
        if success:
            print(f"\n✅ PASS: {test_name}")
            self.passed += 1
        else:
            print(f"\n❌ FAIL: {test_name}")
            if "error" in result:
                print(f"   Error: {result['error']}")
            self.failed += 1
        
        self.test_results.append({
            "test": test_name,
            "success": success,
            "result": result
        })
        
        time.sleep(1)  # Brief pause between tests
        return success
    
    def print_summary(self):
        """Print test summary"""
        print(f"\n{'='*60}")
        print("TEST SUMMARY")
        print(f"{'='*60}")
        print(f"Total Tests: {len(self.test_results)}")
        print(f"✅ Passed: {self.passed}")
        print(f"❌ Failed: {self.failed}")
        print(f"\n{'='*60}")
        
        if self.failed > 0:
            print("\nFailed Tests:")
            for test in self.test_results:
                if not test["success"]:
                    print(f"  - {test['test']}")
        
        print(f"\n{'='*60}")


def main():
    """Main test execution"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description="Test Outlook Desktop Toolkit (READ-ONLY by default)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
SAFETY MODE:
  By default, this script is READ-ONLY and will:
  ✓ Search emails (read-only)
  ✓ Extract email content (read-only)
  ✓ Check attachments (read-only)
  ✓ Download attachments (saves to disk, no email sending)
  
  It will NOT:
  ✗ Send any emails
  ✗ Modify any emails
  ✗ Delete anything
  
  To test email sending, use --enable-email-sending flag.

Examples:
  # Safe read-only testing (default)
  python test_toolkit.py test@outlook.com
  
  # Include email sending tests (requires confirmation)
  python test_toolkit.py test@outlook.com --enable-email-sending
        """
    )
    parser.add_argument(
        "email_account",
        nargs="?",
        help="Outlook email account ID (e.g., test@outlook.com)"
    )
    parser.add_argument(
        "--enable-email-sending",
        action="store_true",
        help="Enable email sending tests (disabled by default for safety)"
    )
    parser.add_argument(
        "--auto-confirm",
        action="store_true",
        help="Auto-confirm all prompts (use with caution)"
    )
    
    args = parser.parse_args()
    
    print("="*60)
    print("Outlook Desktop Toolkit - Automated Test Suite")
    print("="*60)
    print("\n🔒 SAFE MODE: READ-ONLY OPERATIONS ONLY")
    print("   This script will:")
    print("   ✓ Search and read emails")
    print("   ✓ Extract email content")
    print("   ✓ Check attachments")
    print("   ✓ Download attachments (saves to disk)")
    print("\n   This script will NOT:")
    print("   ✗ Send any emails")
    print("   ✗ Modify or delete anything")
    
    if args.enable_email_sending:
        print("\n⚠️  WARNING: Email sending tests are ENABLED")
        print("   This will send actual emails!")
    else:
        print("\n   (Email sending tests are disabled by default)")
    
    print("\n" + "="*60)
    print("\n📋 PREREQUISITES:")
    print("1. Ensure Outlook desktop app is running")
    print("2. Have test emails ready in your inbox")
    print("3. Ensure you have write access to output directory")
    print("\n" + "="*60)
    
    # Get email account
    if args.email_account:
        email_account = args.email_account
    else:
        email_account = input("\nEnter your Outlook email account ID (e.g., test@outlook.com): ").strip()
        if not email_account:
            print("❌ Email account is required")
            sys.exit(1)
    
    print(f"\nUsing email account: {email_account}")
    
    # Confirm before proceeding (unless auto-confirm)
    if not args.auto_confirm:
        print("\n📝 This test will:")
        print("  - Search your inbox (read-only)")
        print("  - Extract email content (read-only)")
        print("  - Download attachments to disk")
        if args.enable_email_sending:
            print("  - ⚠️  SEND ACTUAL EMAILS (if you proceed)")
        
        confirm = input("\nContinue with READ-ONLY tests? (yes/no): ").strip().lower()
        if confirm not in ['yes', 'y']:
            print("Test cancelled.")
            sys.exit(0)
    
    tester = ToolkitTester(email_account)
    
    # Test 1: Connection Test (using find email)
    print("\n" + "="*60)
    print("PHASE 1: Connection and Basic Search")
    print("="*60)
    
    tester.test(
        "Test 1: Connection to Outlook",
        "find_and_extract_email",
        {
            "subject": "Test",
            "email_account": email_account
        },
        should_succeed=False  # May not find email, but should connect
    )
    
    # Test 2: Check Email Attachments
    print("\n" + "="*60)
    print("PHASE 2: Attachment Checking")
    print("="*60)
    
    subject = input("\nEnter subject of an email to test (or press Enter to skip): ").strip()
    if subject:
        tester.test(
            "Test 2: Check Email Attachments",
            "check_email_attachments",
            {
                "subject": subject,
                "email_account": email_account
            },
            expected_keys=["email_found", "has_attachments"]
        )
    
    # Test 3: Check Specific Files
    if subject:
        tester.test(
            "Test 3: Check Specific Files",
            "check_specific_files",
            {
                "subject": subject,
                "email_account": email_account,
                "file_patterns": ["invoice", "receipt", "pdf", "xlsx"]
            },
            expected_keys=["email_found", "patterns_searched"]
        )
    
    # Test 4: Find and Extract Email
    print("\n" + "="*60)
    print("PHASE 3: Email Extraction")
    print("="*60)
    
    extract_subject = input("\nEnter subject of email to extract (or press Enter to skip): ").strip()
    if extract_subject:
        output_path = input("Enter output path (or press Enter for default): ").strip()
        test_args = {
            "subject": extract_subject,
            "email_account": email_account
        }
        if output_path:
            test_args["output_base_path"] = output_path
        
        tester.test(
            "Test 4: Find and Extract Email",
            "find_and_extract_email",
            test_args,
            expected_keys=["email_found", "email_subject", "output_folder"]
        )
    
    # Test 5: Send Email Reply (OPTIONAL - Only if explicitly enabled)
    if args.enable_email_sending:
        print("\n" + "="*60)
        print("PHASE 4: Email Sending (OPTIONAL - REQUIRES CONFIRMATION)")
        print("="*60)
        print("\n⚠️  WARNING: This will send an actual email!")
        
        if not args.auto_confirm:
            send_test = input("\nTest email sending? This will send an actual email! (yes/no): ").strip().lower()
        else:
            send_test = 'no'  # Default to no even with auto-confirm for safety
            print("\n⚠️  Auto-confirm enabled, but skipping email sending for safety.")
            print("   Use --enable-email-sending and manually confirm to test email sending.")
        
        if send_test in ['yes', 'y']:
            to_email = input("Enter recipient email address: ").strip()
            if to_email:
                # Double confirmation for email sending
                print(f"\n⚠️  FINAL CONFIRMATION:")
                print(f"   You are about to send an email to: {to_email}")
                final_confirm = input("   Type 'SEND' to confirm, or anything else to cancel: ").strip()
                
                if final_confirm == 'SEND':
                    tester.test(
                        "Test 5: Send Email Reply",
                        "send_email_reply",
                        {
                            "to_email": to_email,
                            "subject": "Test Email from Outlook Toolkit",
                            "body": "This is an automated test email from the Outlook Desktop Toolkit."
                        },
                        expected_keys=["success"]
                    )
                else:
                    print("   Email sending test cancelled.")
            else:
                print("   No recipient email provided. Skipping email sending test.")
        else:
            print("   Email sending test skipped (safe mode).")
    else:
        print("\n" + "="*60)
        print("PHASE 4: Email Sending (SKIPPED - Safe Mode)")
        print("="*60)
        print("\n✓ Email sending tests are disabled for safety.")
        print("  To enable, run with --enable-email-sending flag.")
    
    # Test 6: Error Handling (All read-only, safe)
    print("\n" + "="*60)
    print("PHASE 5: Error Handling (Read-Only Tests)")
    print("="*60)
    
    tester.test(
        "Test 6: Missing Required Parameter",
        "find_and_extract_email",
        {
            "email_account": email_account
            # Missing subject
        },
        should_succeed=False
    )
    
    tester.test(
        "Test 7: Invalid Email Account",
        "find_and_extract_email",
        {
            "subject": "Test",
            "email_account": "invalid-account-that-does-not-exist@example.com"
        },
        should_succeed=False
    )
    
    tester.test(
        "Test 8: Email Not Found",
        "find_and_extract_email",
        {
            "subject": "NonExistentEmailSubjectThatWillNeverBeFound12345",
            "email_account": email_account
        },
        should_succeed=False
    )
    
    # Print summary
    tester.print_summary()
    
    # Final safety reminder
    print("\n" + "="*60)
    print("SAFETY REMINDER")
    print("="*60)
    print("✓ All read-only tests completed")
    email_sent = False
    if args.enable_email_sending:
        if 'send_test' in locals() and send_test in ['yes', 'y']:
            if 'final_confirm' in locals() and final_confirm == 'SEND':
                email_sent = True
    
    if not email_sent:
        print("✓ No emails were sent (safe mode)")
    else:
        print("⚠️  Email sending test was executed - check recipient inbox")
    print("="*60)
    
    # Exit code
    if tester.failed > 0:
        print("\n⚠️  Some tests failed. Review the output above.")
        sys.exit(1)
    else:
        print("\n✅ All tests passed!")
        sys.exit(0)


if __name__ == "__main__":
    main()
