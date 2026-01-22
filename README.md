# Outlook Desktop Toolkit

A Python toolkit that connects to Microsoft Outlook desktop application to search emails, extract content and attachments, and send automated replies.

## What This Toolkit Does

- **Find emails** by subject (searches your Outlook inbox)
- **Extract email content** and save to text file
- **Download attachments** to organized folders
- **Check if email has attachments** without downloading
- **Check for specific file patterns** in attachments
- **Send automated replies** when attachments are missing (optional)

## Quick Start

### Prerequisites

- Windows computer
- Microsoft Outlook desktop app installed and running
- Python 3.7+ installed

### Installation

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. See [SETUP.md](SETUP.md) for detailed setup instructions

## How to Use

The toolkit accepts JSON input and returns JSON output. It's designed to be called by AI agents or automation systems.

### Basic Usage

**Input Format:**
```json
{
  "capability": "capability_name",
  "args": {
    "parameter1": "value1",
    "parameter2": "value2"
  }
}
```

**Output Format:**
```json
{
  "result": { ... },
  "capability": "capability_name"
}
```
or
```json
{
  "error": "error message",
  "capability": "capability_name"
}
```

### Capability 1: Find and Extract Email

Finds the most recent email matching a subject, extracts content, and downloads attachments.

**Required Parameters:**
- `subject` - Email subject to search for (case-insensitive)
- `email_account` - Your Outlook email account (e.g., "your.email@example.com")

**Optional Parameters:**
- `output_base_path` - Where to save files (default: current directory)
- `search_unread_only` - Search only unread emails (default: true)
- `send_reply_if_no_attachments` - Auto-reply if no attachments (default: false)
- `reply_message` - Custom reply message

**Example:**
```json
{
  "capability": "find_and_extract_email",
  "args": {
    "subject": "Invoice",
    "email_account": "test@outlook.com",
    "output_base_path": "C:/output"
  }
}
```

**Success Response:**
```json
{
  "result": {
    "email_found": true,
    "email_subject": "Invoice #12345",
    "email_sender": "John Doe",
    "has_attachments": true,
    "attachment_count": 2,
    "output_folder": "C:/output/email_extractions/Invoice_12345_2025-01-15_10-30-45",
    "email_content_file": "C:/output/email_extractions/.../email_content.txt",
    "attachments_folder": "C:/output/email_extractions/.../attachments",
    "attachments": [
      {
        "filename": "invoice.pdf",
        "path": "C:/output/.../attachments/invoice.pdf",
        "size_bytes": 123456
      }
    ]
  },
  "capability": "find_and_extract_email"
}
```

### Capability 2: Check Email Attachments

Checks if an email has any attachments without downloading them.

**Required Parameters:**
- `subject` - Email subject to search for
- `email_account` - Your Outlook email account

**Optional Parameters:**
- `search_unread_only` - Search only unread emails (default: true)

**Example:**
```json
{
  "capability": "check_email_attachments",
  "args": {
    "subject": "Invoice",
    "email_account": "test@outlook.com"
  }
}
```

**Success Response:**
```json
{
  "result": {
    "email_found": true,
    "email_subject": "Invoice #12345",
    "email_sender": "John Doe",
    "has_attachments": true,
    "attachment_count": 2,
    "attachments": [
      {
        "filename": "invoice.pdf",
        "size_bytes": 123456
      },
      {
        "filename": "receipt.pdf",
        "size_bytes": 78901
      }
    ]
  },
  "capability": "check_email_attachments"
}
```

### Capability 3: Check Specific Files

Checks if an email contains specific file patterns in its attachments.

**Required Parameters:**
- `subject` - Email subject to search for
- `email_account` - Your Outlook email account
- `file_patterns` - List of file name patterns to search for (e.g., `["invoice", "receipt", "contract"]`)

**Optional Parameters:**
- `search_unread_only` - Search only unread emails (default: true)

**Example:**
```json
{
  "capability": "check_specific_files",
  "args": {
    "subject": "Invoice",
    "email_account": "test@outlook.com",
    "file_patterns": ["invoice", "receipt", "contract", "statement"]
  }
}
```

**Success Response:**
```json
{
  "result": {
    "email_found": true,
    "email_subject": "Invoice #12345",
    "email_sender": "John Doe",
    "has_attachments": true,
    "attachment_count": 3,
    "all_attachments": ["invoice.pdf", "receipt.pdf", "notes.txt"],
    "patterns_searched": ["invoice", "receipt", "contract", "statement"],
    "found_patterns": ["invoice", "receipt"],
    "missing_patterns": ["contract", "statement"],
    "pattern_details": {
      "invoice": {
        "found": true,
        "matching_files": ["invoice.pdf"]
      },
      "receipt": {
        "found": true,
        "matching_files": ["receipt.pdf"]
      },
      "contract": {
        "found": false,
        "matching_files": []
      },
      "statement": {
        "found": false,
        "matching_files": []
      }
    },
    "all_patterns_found": false
  },
  "capability": "check_specific_files"
}
```

### Capability 4: Send Email Reply

Sends an email reply via Outlook.

**Required Parameters:**
- `to_email` - Recipient email address
- `subject` - Email subject
- `body` - Email body content

**Optional Parameters:**
- `email_account` - Account to send from (uses default if not specified)

**Example:**
```json
{
  "capability": "send_email_reply",
  "args": {
    "to_email": "sender@example.com",
    "subject": "Re: Your Request",
    "body": "Thank you for your email."
  }
}
```

## Configuration: What to Set Up Front vs. Runtime

### Set Up Once (Configuration)

These are typically configured once during setup:

1. **Email Account ID** - Your Outlook account identifier
   - Find it using the method in [SETUP.md](SETUP.md)
   - Can be hardcoded in your calling application or passed each time

2. **Default Output Path** - Where files are saved
   - Can be set in `config.py` as `DEFAULT_OUTPUT_BASE_PATH`
   - Or specified per request via `output_base_path` parameter

3. **Default Reply Message** - Template for automated replies
   - Can be set in `config.py` as `DEFAULT_REPLY_MESSAGE`
   - Or specified per request via `reply_message` parameter

### Provided at Runtime (Via Prompt/Request)

These are provided each time the toolkit is called:

- `subject` - What email to search for
- `output_base_path` - Where to save (if different from default)
- `search_unread_only` - Search behavior
- `send_reply_if_no_attachments` - Whether to auto-reply
- `reply_message` - Custom reply text
- `to_email`, `subject`, `body` - For sending replies

## Example AI Agent Prompts

Here are example prompts users might use when calling this toolkit through an AI agent:

### Example 1: Find Email with Attachments
```
"Find the most recent email with subject containing 'Invoice' in my Outlook inbox (test@outlook.com) and download all attachments to C:/downloads"
```

**Agent would call:**
```json
{
  "capability": "find_and_extract_email",
  "args": {
    "subject": "Invoice",
    "email_account": "test@outlook.com",
    "output_base_path": "C:/downloads"
  }
}
```

### Example 2: Check for Email and Auto-Reply if No Attachments
```
"Look for an email about 'Report' in my unread emails. If it has no attachments, send a reply asking for the files."
```

**Agent would call:**
```json
{
  "capability": "find_and_extract_email",
  "args": {
    "subject": "Report",
    "email_account": "test@outlook.com",
    "search_unread_only": true,
    "send_reply_if_no_attachments": true
  }
}
```

### Example 3: Extract Email Content Only
```
"Find the email with 'Meeting Notes' in the subject and save its content to a text file"
```

**Agent would call:**
```json
{
  "capability": "find_and_extract_email",
  "args": {
    "subject": "Meeting Notes",
    "email_account": "test@outlook.com"
  }
}
```

### Example 4: Send Custom Reply
```
"Send an email to john@example.com with subject 'Re: Your Request' and body 'We received your request and will process it shortly.'"
```

**Agent would call:**
```json
{
  "capability": "send_email_reply",
  "args": {
    "to_email": "john@example.com",
    "subject": "Re: Your Request",
    "body": "We received your request and will process it shortly."
  }
}
```

### Example 5: Check if Email Has Attachments
```
"Check if the email with 'Invoice' in the subject has any attachments"
```

**Agent would call:**
```json
{
  "capability": "check_email_attachments",
  "args": {
    "subject": "Invoice",
    "email_account": "test@outlook.com"
  }
}
```

### Example 6: Check for Specific File Types
```
"Check if the email about 'Report' has files containing 'invoice', 'receipt', 'contract', or 'statement' in their names"
```

**Agent would call:**
```json
{
  "capability": "check_specific_files",
  "args": {
    "subject": "Report",
    "email_account": "test@outlook.com",
    "file_patterns": ["invoice", "receipt", "contract", "statement"]
  }
}
```

### Example 7: Search All Emails (Including Read)
```
"Find any email (read or unread) with 'Contract' in the subject and extract everything"
```

**Agent would call:**
```json
{
  "capability": "find_and_extract_email",
  "args": {
    "subject": "Contract",
    "email_account": "test@outlook.com",
    "search_unread_only": false
  }
}
```

## Output Folder Structure

When an email is extracted, files are organized like this:

```
{output_base_path}/
└── email_extractions/
    └── {subject}_{timestamp}/
        ├── email_content.txt    (email metadata and body)
        └── attachments/
            ├── file1.pdf
            └── file2.xlsx
```

## Error Responses

If something goes wrong, you'll get an error response:

```json
{
  "error": "No email found with subject containing: Invoice",
  "capability": "find_and_extract_email"
}
```

Common errors:
- `"No email found with subject containing: ..."` - No matching email
- `"Failed to connect to Outlook"` - Outlook not running or not installed
- `"Could not find folder for account: ..."` - Wrong email account ID
- `"Missing required parameter: ..."` - Required parameter not provided

## Testing

Test the toolkit locally:

```bash
# Test finding email
echo '{"capability": "find_and_extract_email", "args": {"subject": "Test", "email_account": "your.email@example.com"}}' | python main.py

# Test sending reply
echo '{"capability": "send_email_reply", "args": {"to_email": "test@example.com", "subject": "Test", "body": "Test message"}}' | python main.py
```

## Documentation

- **[SETUP.md](SETUP.md)** - Detailed setup and configuration guide
- **toolkit.json** - Toolkit metadata and capability schemas

## Important Notes

- **Windows only** - Requires Windows and Outlook desktop app
- **Outlook must be running** - The toolkit connects to running Outlook
- **Case-insensitive search** - Subject search matches any case
- **Most recent first** - Returns the newest matching email
- **Inbox only** - Searches only the Inbox folder (not subfolders)

## Support

For setup help, troubleshooting, and detailed configuration, see [SETUP.md](SETUP.md).
