# Outlook Desktop Toolkit - Setup Guide

This document provides a comprehensive step-by-step guide for setting up and configuring the Outlook Desktop Toolkit.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation Steps](#installation-steps)
3. [Outlook Configuration](#outlook-configuration)
4. [Finding Your Email Account ID](#finding-your-email-account-id)
5. [Configuration Options](#configuration-options)
6. [Testing Your Setup](#testing-your-setup)
7. [Troubleshooting Setup Issues](#troubleshooting-setup-issues)

---

## Prerequisites

### 1. System Requirements

- **Operating System**: Windows 7 or later (Windows 10/11 recommended)
- **Python**: Version 3.7 or higher
- **Microsoft Outlook**: Desktop application (not Outlook Web Access)
- **Administrator Rights**: May be required for initial installation

### 2. Verify Python Installation

Open Command Prompt or PowerShell and run:

```bash
python --version
```

Expected output: `Python 3.7.x` or higher

If Python is not installed:
1. Download from [python.org](https://www.python.org/downloads/)
2. During installation, check "Add Python to PATH"
3. Verify installation after completion

### 3. Verify Outlook Installation

1. Open Microsoft Outlook desktop application
2. Ensure it launches without errors
3. Verify you can send and receive emails
4. Check that at least one email account is configured

**Note**: Outlook Web Access (OWA) or Outlook.com web interface will NOT work. You need the desktop application.

---

## Installation Steps

### Step 1: Navigate to Toolkit Directory

Open Command Prompt or PowerShell and navigate to the toolkit folder:

```bash
cd path\to\Email-utility\Outlook_desktop_toolkit
```

### Step 2: Install Dependencies

Install required Python packages:

```bash
pip install -r requirements.txt
```

This installs:
- `pywin32` (version 306 or higher) - Windows COM API interface

**Expected output:**
```
Successfully installed pywin32-306 ...
```

### Step 3: Verify Installation

Test that pywin32 is properly installed:

```bash
python -c "import win32com.client; print('✓ COM API available')"
```

If you see "✓ COM API available", the installation is successful.

### Step 4: Test Outlook Connection

Create a test script to verify Outlook connectivity:

**Create file: `test_outlook_connection.py`**
```python
import win32com.client

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    print("✓ Successfully connected to Outlook")
    
    # List available accounts
    print("\nAvailable email accounts:")
    for folder in namespace.Folders:
        print(f"  - {folder.Name}")
        
except Exception as e:
    print(f"✗ Error connecting to Outlook: {e}")
    print("\nTroubleshooting:")
    print("1. Ensure Outlook desktop application is installed")
    print("2. Ensure Outlook is running")
    print("3. Ensure at least one email account is configured")
```

Run the test:
```bash
python test_outlook_connection.py
```

**Expected output:**
```
✓ Successfully connected to Outlook

Available email accounts:
  - test@gmail.com
  - another@example.com
```

---

## Outlook Configuration

### 1. Launch Outlook

Ensure Microsoft Outlook desktop application is running before using the toolkit.

### 2. Configure Email Accounts

1. Open Outlook
2. Go to **File → Account Settings → Account Settings**
3. Verify your email accounts are listed and configured correctly
4. Test sending/receiving emails to ensure accounts are working

### 3. Outlook Security Settings

The toolkit uses COM automation, which may trigger Outlook security prompts:

**For Outlook 2010 and later:**
- First-time use may show a security warning
- Click "Allow" or "Allow access for X minutes"
- To avoid repeated prompts, you can configure Outlook security settings (requires registry changes - advanced users only)

**Note**: If you see security prompts, this is normal. Allow the access to proceed.

---

## Finding Your Email Account ID

The `email_account` parameter must match exactly how Outlook identifies your account. Here are methods to find it:

### Method 1: Using Outlook UI

1. Open Outlook
2. Go to **File → Account Settings → Account Settings**
3. Click on the **Email** tab
4. Look at the account name/email address
5. This is your `email_account` ID

**Common formats:**
- Email address: `your.email@example.com`
- Display name: `Your Name`
- SMTP address: `your.email@example.com`

### Method 2: Using Python Script

Use the test script from Step 4 above, or create this simple script:

```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

print("Available email accounts (use these as email_account parameter):")
for folder in namespace.Folders:
    print(f"  '{folder.Name}'")
    
    # Try to get inbox
    try:
        inbox = folder.Folders("Inbox")
        print(f"    → Inbox found with {inbox.Items.Count} items")
    except:
        print(f"    → Inbox not accessible")
```

### Method 3: Trial and Error

If unsure, try these in order:
1. Your email address: `your.email@example.com`
2. Your display name as shown in Outlook
3. Run the Python script above to see exact names

**Important**: The account ID is case-sensitive and must match exactly.

---

## Configuration Options

### 1. Default Settings (config.py)

Edit `config.py` to change default behavior:

```python
class ToolkitConfig:
    # Default output directory
    DEFAULT_OUTPUT_BASE_PATH = os.getcwd()  # Current directory
    
    # Default reply message template
    DEFAULT_REPLY_MESSAGE = "Please provide the required attachments for: {subject}"
    
    # Default search behavior
    DEFAULT_SEARCH_UNREAD_ONLY = True
    
    # Default reply behavior
    DEFAULT_SEND_REPLY_IF_NO_ATTACHMENTS = False
```

### 2. Per-Request Configuration

You can override defaults in each JSON request:

```json
{
  "capability": "find_and_extract_email",
  "args": {
    "subject": "Invoice",
    "email_account": "test@gmail.com",
    "output_base_path": "C:/MyOutput",           // Override default path
    "search_unread_only": false,                 // Override default search
    "send_reply_if_no_attachments": true,        // Override default reply
    "reply_message": "Custom message: {subject}" // Override default message
  }
}
```

### 3. Output Directory Configuration

**Option A: Use default (current directory)**
- No configuration needed
- Files saved to: `./email_extractions/`

**Option B: Set in config.py**
```python
DEFAULT_OUTPUT_BASE_PATH = r"C:\EmailExtractions"
```

**Option C: Specify per request**
```json
"output_base_path": "D:/MyEmailData"
```

**Directory Structure Created:**
```
{output_base_path}/
└── email_extractions/
    └── {subject}_{timestamp}/
        ├── email_content.txt
        └── attachments/
            └── [attachment files]
```

### 4. Reply Message Templates

**Default Template:**
```
"Please provide the required attachments for: {subject}"
```

**Custom Template Example:**
```json
"reply_message": "Hello,\n\nWe received your email about '{subject}' but noticed no attachments were included.\n\nPlease resend with the required files.\n\nThank you!"
```

The `{subject}` placeholder will be replaced with the actual email subject.

---

## Testing Your Setup

### Test 1: Basic Connection Test

```bash
python test_outlook_connection.py
```

**Expected**: List of available email accounts

### Test 2: Find Email Test

```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test", "email_account": "your.email@example.com"}}' | python main.py
```

**Expected**: JSON response with email details or error message

### Test 3: Send Reply Test

```bash
echo '{"capability": "send_email_reply", "args": {"to_email": "test@example.com", "subject": "Test", "body": "This is a test"}}' | python main.py
```

**Expected**: JSON response with success confirmation

### Test 4: Full Workflow Test

1. Send yourself an email with subject "Test Email" and an attachment
2. Run the find_and_extract_email capability
3. Verify:
   - Email is found
   - Content file is created
   - Attachment is downloaded
   - Output paths are correct

---

## Troubleshooting Setup Issues

### Issue 1: "ModuleNotFoundError: No module named 'win32com'"

**Cause**: pywin32 not installed or not properly installed

**Solution**:
```bash
pip uninstall pywin32
pip install pywin32
```

If still failing, try:
```bash
python -m pip install --upgrade pywin32
```

### Issue 2: "Failed to connect to Outlook"

**Possible Causes & Solutions**:

1. **Outlook not running**
   - Solution: Launch Microsoft Outlook desktop application
   - Verify: Outlook icon appears in system tray

2. **Outlook not installed**
   - Solution: Install Microsoft Outlook desktop application
   - Note: Office 365 subscription includes Outlook

3. **No email accounts configured**
   - Solution: Configure at least one email account in Outlook
   - Verify: File → Account Settings → Account Settings shows accounts

4. **COM registration issues**
   - Solution: Re-register COM components:
     ```bash
     python Scripts/pywin32_postinstall.py -install
     ```
   - Or run as administrator:
     ```bash
     python -m win32com.client.makepy
     ```

### Issue 3: "Could not find folder for account: ..."

**Cause**: Email account ID doesn't match

**Solution**:
1. Run the account listing script to see exact account names
2. Use the exact name shown (case-sensitive)
3. Common mistakes:
   - Using display name instead of email address
   - Case mismatch
   - Extra spaces

### Issue 4: Security Warnings from Outlook

**Cause**: Outlook security settings blocking COM access

**Solution**:
1. When prompted, click "Allow" or "Allow access for X minutes"
2. For persistent access (advanced):
   - Modify Outlook security settings (requires registry changes)
   - Or use Outlook's Trust Center settings

### Issue 5: "Permission denied" when saving files

**Cause**: Insufficient permissions on output directory

**Solution**:
1. Choose a directory you have write access to
2. Run Command Prompt as Administrator if needed
3. Check folder permissions in Windows Explorer

### Issue 6: Python not found

**Cause**: Python not in PATH

**Solution**:
1. Reinstall Python with "Add to PATH" option
2. Or use full path: `C:\Python39\python.exe`
3. Or add Python to PATH manually in System Environment Variables

---

## Quick Reference

### Required Parameters

**find_and_extract_email:**
- `subject` (string): Email subject to search
- `email_account` (string): Outlook account ID

**send_email_reply:**
- `to_email` (string): Recipient address
- `subject` (string): Email subject
- `body` (string): Email body

### Optional Parameters

**find_and_extract_email:**
- `output_base_path` (string): Output directory
- `search_unread_only` (boolean): Search only unread (default: true)
- `send_reply_if_no_attachments` (boolean): Auto-reply if no attachments (default: false)
- `reply_message` (string): Custom reply message

**send_email_reply:**
- `email_account` (string): Account to send from (uses default if not specified)

### Common Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Test connection
python -c "import win32com.client; print('OK')"

# Run toolkit
echo '{"capability": "...", "args": {...}}' | python main.py
```

---

## Next Steps

After setup is complete:

1. Read the main [README.md](README.md) for usage examples
2. Test with your own emails
3. Configure default settings in `config.py` if needed
4. Integrate into your automation workflow

For additional help, refer to the Troubleshooting section in [README.md](README.md).
