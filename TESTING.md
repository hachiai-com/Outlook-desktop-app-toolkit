# Outlook Desktop Toolkit - Testing Guide

This guide helps you test the Outlook Desktop Toolkit on a client machine to verify it works correctly with the desktop Outlook application.

## Prerequisites for Testing

1. **Windows machine** with Outlook desktop app installed
2. **Outlook must be running** - Launch Outlook before testing
3. **At least one email account** configured in Outlook
4. **Test emails in inbox** - Have some emails ready for testing
5. **Python 3.7+** installed and accessible

## Quick Test Checklist

- [ ] Outlook desktop app is running
- [ ] Python can import required modules
- [ ] Toolkit can connect to Outlook
- [ ] Can find email by subject
- [ ] Can extract email content
- [ ] Can download attachments
- [ ] Can check for attachments
- [ ] Can check for specific files
- [ ] Can send email replies

## Step-by-Step Testing

### Step 1: Verify Prerequisites

#### 1.1 Check Python Installation
```bash
python --version
```
Expected: Python 3.7.x or higher

#### 1.2 Check Dependencies
```bash
python -c "import win32com.client; print('✓ pywin32 installed')"
```
Expected: `✓ pywin32 installed`

#### 1.3 Verify Outlook is Running
- Open Microsoft Outlook desktop application
- Verify it's not minimized or closed
- Check system tray for Outlook icon

### Step 2: Find Your Email Account ID

Before testing, you need to know your email account ID. Run:

```bash
python find_account_id.py
```

Or use this quick script:
```python
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
print("Available email accounts:")
for folder in namespace.Folders:
    print(f"  - {folder.Name}")
```

**Save the account ID** - you'll need it for all tests.

### Step 3: Prepare Test Emails

Before testing, ensure you have:

1. **At least one unread email** with a unique subject (e.g., "Test Email 2025-01-15")
2. **One email with attachments** (e.g., "Test Invoice with Attachments")
3. **One email without attachments** (e.g., "Test Email No Attachments")

**Tip:** Send yourself test emails with known subjects and attachments.

### Step 4: Run Automated Test Script

Use the provided test script for comprehensive testing:

```bash
# Safe read-only testing (default - NO emails will be sent)
python test_toolkit.py your-email@example.com

# Or run interactively (will prompt for email account)
python test_toolkit.py
```

**🔒 SAFETY: By default, the test script is READ-ONLY and will NOT send emails.**

This will test all capabilities automatically. See [Automated Testing](#automated-testing) section below.

### Step 5: Manual Testing (Alternative)

If you prefer manual testing, follow the sections below.

## Manual Testing Procedures

### Test 1: Connection Test

**Purpose:** Verify toolkit can connect to Outlook

**Command:**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test", "email_account": "YOUR_EMAIL@example.com"}}' | python main.py
```

**Expected Result:**
- If email found: JSON with `"email_found": true`
- If no email: JSON with `"error": "No email found..."`
- If connection fails: JSON with `"error": "Failed to connect to Outlook"`

**Success Criteria:**
- No "Failed to connect to Outlook" error
- Toolkit responds with valid JSON

### Test 2: Find Email by Subject

**Purpose:** Verify email search functionality

**Command:**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test Email", "email_account": "YOUR_EMAIL@example.com"}}' | python main.py
```

**Expected Result:**
```json
{
  "result": {
    "email_found": true,
    "email_subject": "Test Email 2025-01-15",
    "email_sender": "Your Name",
    "has_attachments": true/false,
    ...
  },
  "capability": "find_and_extract_email"
}
```

**Success Criteria:**
- Finds the most recent email matching subject
- Returns correct email metadata
- No errors

### Test 3: Extract Email Content

**Purpose:** Verify email content extraction and file saving

**Command:**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test Email", "email_account": "YOUR_EMAIL@example.com", "output_base_path": "C:/test_output"}}' | python main.py
```

**Expected Result:**
- Email content saved to text file
- Output folder created with timestamp
- JSON response includes `email_content_file` path

**Verify:**
1. Check output folder exists: `C:/test_output/email_extractions/...`
2. Open `email_content.txt` - should contain email body and metadata
3. Verify content is readable and complete

**Success Criteria:**
- File created successfully
- Content matches email
- Metadata is correct

### Test 4: Download Attachments

**Purpose:** Verify attachment downloading

**Prerequisites:** Email with attachments in inbox

**Command:**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test Invoice", "email_account": "YOUR_EMAIL@example.com", "output_base_path": "C:/test_output"}}' | python main.py
```

**Expected Result:**
- Attachments downloaded to `attachments` folder
- JSON response includes attachment list with paths

**Verify:**
1. Check `attachments` folder exists
2. Verify all attachment files are present
3. Open files to ensure they're not corrupted
4. Check file sizes match

**Success Criteria:**
- All attachments downloaded
- Files are accessible
- File sizes correct

### Test 5: Check Email Attachments (No Download)

**Purpose:** Verify attachment checking without downloading

**Command:**
```bash
echo '{"capability": "check_email_attachments", "args": {"subject": "Test Email", "email_account": "YOUR_EMAIL@example.com"}}' | python main.py
```

**Expected Result:**
```json
{
  "result": {
    "email_found": true,
    "has_attachments": true,
    "attachment_count": 2,
    "attachments": [...]
  },
  "capability": "check_email_attachments"
}
```

**Success Criteria:**
- Returns correct attachment count
- Lists attachment filenames and sizes
- No files downloaded (check output folder)

### Test 6: Check Specific Files

**Purpose:** Verify file pattern matching

**Prerequisites:** Email with attachments containing specific patterns

**Command:**
```bash
echo '{"capability": "check_specific_files", "args": {"subject": "Test Documents", "email_account": "YOUR_EMAIL@example.com", "file_patterns": ["invoice", "receipt", "contract"]}}' | python main.py
```

**Expected Result:**
```json
{
  "result": {
    "email_found": true,
    "found_patterns": ["invoice", "receipt"],
    "missing_patterns": ["contract"],
    "pattern_details": {...}
  },
  "capability": "check_specific_files"
}
```

**Success Criteria:**
- Correctly identifies matching files
- Reports missing patterns
- Pattern matching is case-insensitive

### Test 7: Search Unread vs All Emails

**Purpose:** Verify search scope options

**Test 7a: Search Unread Only (Default)**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test", "email_account": "YOUR_EMAIL@example.com", "search_unread_only": true}}' | python main.py
```

**Test 7b: Search All Emails**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test", "email_account": "YOUR_EMAIL@example.com", "search_unread_only": false}}' | python main.py
```

**Success Criteria:**
- `search_unread_only: true` only finds unread emails
- `search_unread_only: false` finds both read and unread

### Test 8: Send Email Reply

**Purpose:** Verify email sending functionality

**⚠️ WARNING:** This will send an actual email!

**Command:**
```bash
echo '{"capability": "send_email_reply", "args": {"to_email": "your-test-email@example.com", "subject": "Test Email from Toolkit", "body": "This is a test email from the Outlook Desktop Toolkit."}}' | python main.py
```

**Expected Result:**
```json
{
  "result": {
    "success": true,
    "to": "your-test-email@example.com",
    "subject": "Test Email from Toolkit",
    "message": "Email sent successfully"
  },
  "capability": "send_email_reply"
}
```

**Verify:**
1. Check recipient's inbox for the email
2. Verify email content is correct
3. Check sender account is correct

**Success Criteria:**
- Email sent successfully
- Recipient receives email
- Content matches request

### Test 9: Auto-Reply on Missing Attachments

**Purpose:** Verify auto-reply functionality

**Prerequisites:** Email without attachments

**Command:**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test No Attachments", "email_account": "YOUR_EMAIL@example.com", "send_reply_if_no_attachments": true, "reply_message": "Please send the attachments for: {subject}"}}' | python main.py
```

**Expected Result:**
- Email found with no attachments
- Reply sent automatically
- JSON response includes `reply_sent: true`

**Verify:**
1. Check original sender's inbox for reply
2. Verify reply message content

**Success Criteria:**
- Reply sent when no attachments
- Reply message is correct
- Original sender receives reply

### Test 10: Error Handling

**Purpose:** Verify error handling works correctly

**Test 10a: Invalid Email Account**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test", "email_account": "invalid@account.com"}}' | python main.py
```
Expected: Error about account not found

**Test 10b: Missing Subject**
```bash
echo '{"capability": "find_and_extract_email", "args": {"email_account": "YOUR_EMAIL@example.com"}}' | python main.py
```
Expected: Error about missing subject parameter

**Test 10c: Email Not Found**
```bash
echo '{"capability": "find_and_extract_email", "args": {"subject": "NonExistentEmailSubject12345", "email_account": "YOUR_EMAIL@example.com"}}' | python main.py
```
Expected: Error about email not found

**Success Criteria:**
- Errors are returned in JSON format
- Error messages are clear and helpful
- Toolkit doesn't crash

## Automated Testing

The `test_toolkit.py` script provides automated testing with built-in safety features.

### Safe Mode (Default - Recommended)

**By default, the script is READ-ONLY and will NOT send emails:**

```bash
# Run with email account as argument
python test_toolkit.py your-email@example.com

# Or run interactively (will prompt for email account)
python test_toolkit.py
```

**What it does (READ-ONLY):**
- ✅ Searches emails (read-only)
- ✅ Extracts email content (read-only)
- ✅ Checks attachments (read-only)
- ✅ Downloads attachments (saves to disk)
- ❌ **Does NOT send emails**
- ❌ **Does NOT modify or delete anything**

### Testing Email Sending (Optional)

If you need to test email sending functionality, use the `--enable-email-sending` flag:

```bash
python test_toolkit.py your-email@example.com --enable-email-sending
```

**⚠️ WARNING:** This will send actual emails and requires multiple confirmations.

**Safety features:**
- Email sending is disabled by default
- Requires explicit `--enable-email-sending` flag
- Requires user confirmation before sending
- Requires typing 'SEND' as final confirmation
- Shows clear warnings before any email operation

### Command Line Options

```bash
# Show help
python test_toolkit.py --help

# Safe read-only testing (default)
python test_toolkit.py your-email@example.com

# Enable email sending tests (requires confirmation)
python test_toolkit.py your-email@example.com --enable-email-sending

# Auto-confirm prompts (use with caution, still won't send emails by default)
python test_toolkit.py your-email@example.com --auto-confirm
```

## Troubleshooting Test Failures

### Issue: "Failed to connect to Outlook"
**Solution:**
- Ensure Outlook desktop app is running
- Check Outlook is not in "safe mode"
- Restart Outlook and try again

### Issue: "Could not find folder for account"
**Solution:**
- Verify email account ID is correct (case-sensitive)
- Run `find_account_id.py` to see exact account names
- Check account is configured in Outlook

### Issue: "No email found"
**Solution:**
- Verify email exists in Inbox (not subfolders)
- Check subject search term matches (case-insensitive)
- Try `search_unread_only: false` to search all emails
- Verify email hasn't been moved or deleted

### Issue: "Permission denied" when saving files
**Solution:**
- Check output path permissions
- Use a path you have write access to
- Run Command Prompt as Administrator if needed

### Issue: Email sent but not received
**Solution:**
- Check recipient's spam folder
- Verify recipient email address is correct
- Check Outlook send/receive status
- Verify email account has send permissions

## Test Results Template

Document your test results:

```
Test Date: ___________
Tester: ___________
Outlook Version: ___________
Python Version: ___________

Test Results:
[ ] Test 1: Connection - PASS / FAIL
[ ] Test 2: Find Email - PASS / FAIL
[ ] Test 3: Extract Content - PASS / FAIL
[ ] Test 4: Download Attachments - PASS / FAIL
[ ] Test 5: Check Attachments - PASS / FAIL
[ ] Test 6: Check Specific Files - PASS / FAIL
[ ] Test 7: Search Scope - PASS / FAIL
[ ] Test 8: Send Email - PASS / FAIL
[ ] Test 9: Auto-Reply - PASS / FAIL
[ ] Test 10: Error Handling - PASS / FAIL

Notes:
_______________________________________
_______________________________________
```

## Next Steps

After successful testing:
1. Document any issues found
2. Verify all capabilities work as expected
3. Test with real-world scenarios
4. Prepare for integration with consuming application

For integration help, see [README.md](README.md) and [agent_detail.md](agent_detail.md).
