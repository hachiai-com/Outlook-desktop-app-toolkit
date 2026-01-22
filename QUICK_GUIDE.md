# Quick Guide - Using Outlook Desktop Toolkit

A simple guide for non-technical users on how to use this toolkit, especially when called by AI agents.

## What You Need to Know

### Before First Use (One-Time Setup)

1. **Your Outlook Email Account ID**
   - This is usually your email address (e.g., `your.email@example.com`)
   - See [SETUP.md](SETUP.md) for how to find it exactly
   - You'll need this for every request

2. **Where to Save Files** (Optional)
   - Default: Files save to current directory
   - You can specify a different location each time

### When Making Requests

You need to provide:
- **What to search for**: The email subject (or part of it)
- **Which account**: Your Outlook email account ID
- **Where to save**: Output directory (optional)

## Common Use Cases

### Use Case 1: "Find and Download Email Attachments"

**What you want:** Find an email and get its attachments

**What to provide:**
- Subject to search for (e.g., "Invoice", "Report", "Contract")
- Your email account
- Where to save files (optional)

**Example prompt for AI agent:**
```
"Find the email with 'Invoice' in the subject from my Outlook account test@gmail.com and download all attachments"
```

### Use Case 2: "Get Email Content as Text File"

**What you want:** Save email content to a text file

**What to provide:**
- Subject to search for
- Your email account
- Where to save (optional)

**Example prompt:**
```
"Find the email about 'Meeting Notes' and save its content to a text file"
```

### Use Case 3: "Check Email and Reply if No Attachments"

**What you want:** Find email, and if it has no attachments, automatically reply asking for files

**What to provide:**
- Subject to search for
- Your email account
- Option to enable auto-reply

**Example prompt:**
```
"Look for an email with 'Report' in the subject. If it has no attachments, send a reply asking for the files"
```

### Use Case 4: "Send a Reply Email"

**What you want:** Send an email reply

**What to provide:**
- Recipient email address
- Subject
- Message body

**Example prompt:**
```
"Send an email to john@example.com with subject 'Re: Your Request' saying 'Thank you, we received your request'"
```

## Understanding the Response

### Success Response

When successful, you'll get:
- Email details (subject, sender, date)
- File locations (where content and attachments were saved)
- Attachment list with file names and sizes

### Error Response

If something fails, you'll get:
- Error message explaining what went wrong
- Common issues:
  - Email not found → Try different subject or check if email exists
  - Outlook not running → Start Outlook desktop app
  - Wrong account → Check your email account ID

## Tips for Best Results

1. **Be specific with subject**: Use unique words from the email subject
2. **Check Outlook is running**: The toolkit needs Outlook desktop app to be open
3. **Use correct account ID**: Must match exactly (see SETUP.md)
4. **Check output folder**: Files are saved in organized folders with timestamps

## What Gets Created

When an email is processed:
- A folder named after the email subject and timestamp
- A text file with email content (`email_content.txt`)
- An `attachments` folder with all downloaded files

Example:
```
email_extractions/
└── Invoice_12345_2025-01-15_10-30-45/
    ├── email_content.txt
    └── attachments/
        ├── invoice.pdf
        └── receipt.pdf
```

## Need More Help?

- **Setup issues?** → See [SETUP.md](SETUP.md)
- **Technical details?** → See [README.md](README.md)
- **How it works?** → See [README.md](README.md) for capability details
