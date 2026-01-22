#!/usr/bin/env python3
"""
Main entry point for Outlook Desktop Toolkit
Reads JSON from stdin and outputs JSON to stdout
"""
import json
import sys
import logging
from typing import Dict, Any
from email_processor import EmailProcessor
from email_sender import EmailSender
from outlook_connector import OutlookConnector

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def find_and_extract_email(args: Dict[str, Any]) -> Dict[str, Any]:
    """
    Find email by subject and extract content and attachments
    
    Args:
        args: Dictionary with capability arguments
        
    Returns:
        Dictionary with result or error
    """
    try:
        # Validate required parameters
        subject = args.get("subject")
        email_account = args.get("email_account")
        
        if not subject:
            return {
                "error": "Missing required parameter: subject",
                "capability": "find_and_extract_email"
            }
        
        if not email_account:
            return {
                "error": "Missing required parameter: email_account",
                "capability": "find_and_extract_email"
            }
        
        # Get optional parameters
        output_base_path = args.get("output_base_path")
        search_unread_only = args.get("search_unread_only", True)
        send_reply_if_no_attachments = args.get("send_reply_if_no_attachments", False)
        reply_message = args.get("reply_message")
        
        # Process email
        processor = EmailProcessor()
        result = processor.process_email(
            subject=subject,
            email_account=email_account,
            output_base_path=output_base_path,
            search_unread_only=search_unread_only
        )
        
        # Check if email was found
        if not result.get("email_found", False):
            return {
                "error": result.get("error", "Email not found"),
                "capability": "find_and_extract_email"
            }
        
        # Handle case when no attachments and reply is requested
        if not result.get("has_attachments", False) and send_reply_if_no_attachments:
            connector = None
            try:
                # Get the original email to send reply
                connector = OutlookConnector()
                connector.initialize_com()
                inbox, _ = connector.get_inbox(email_account)
                
                # Find the email again to get the object for reply
                if search_unread_only:
                    items = inbox.Items.Restrict("[Unread] = True")
                else:
                    items = inbox.Items
                
                items.Sort("[ReceivedTime]", True)
                subject_lower = subject.lower()
                email = items.GetFirst()
                
                original_email = None
                while email:
                    try:
                        if email.Class == 43:  # MailItem
                            email_subject = getattr(email, 'Subject', '')
                            if subject_lower in email_subject.lower():
                                original_email = email
                                break
                    except Exception:
                        pass
                    email = items.GetNext()
                
                # Send reply if email found (extract properties while COM is active)
                if original_email:
                    # Extract email properties while COM is still active
                    sender_email = getattr(original_email, 'SenderEmailAddress', None)
                    if not sender_email:
                        sender_email = getattr(original_email, 'Sender', None)
                        if hasattr(sender_email, 'Address'):
                            sender_email = sender_email.Address
                    
                    original_subject = getattr(original_email, 'Subject', 'Your email')
                    
                    # Now uninitialize COM - we have all the data we need
                    connector.uninitialize_com()
                    connector = None
                    
                    # Send reply using extracted data
                    sender = EmailSender()
                    reply_subject = f"Re: {original_subject}"
                    
                    if reply_message is None:
                        from config import ToolkitConfig
                        config = ToolkitConfig()
                        reply_message = config.DEFAULT_REPLY_MESSAGE.format(
                            subject=original_subject
                        )
                    
                    reply_result = sender.send_reply(
                        to_email=sender_email,
                        subject=reply_subject,
                        body=reply_message,
                        email_account=email_account
                    )
                    
                    result["reply_sent"] = reply_result.get("success", False)
                    if not reply_result.get("success", False):
                        result["reply_error"] = reply_result.get("error", "Unknown error")
                else:
                    if connector:
                        connector.uninitialize_com()
                    logger.warning("Could not find email object to send reply")
                    result["reply_sent"] = False
                    result["reply_error"] = "Could not find email object for reply"
                    
            except Exception as e:
                logger.error(f"Error sending reply: {str(e)}")
                if connector:
                    try:
                        connector.uninitialize_com()
                    except:
                        pass
                result["reply_sent"] = False
                result["reply_error"] = str(e)
        
        return {
            "result": result,
            "capability": "find_and_extract_email"
        }
        
    except Exception as e:
        logger.error(f"Error in find_and_extract_email: {str(e)}")
        return {
            "error": str(e),
            "capability": "find_and_extract_email"
        }


def check_email_attachments(args: Dict[str, Any]) -> Dict[str, Any]:
    """
    Check if email has any attachments
    
    Args:
        args: Dictionary with capability arguments
        
    Returns:
        Dictionary with result or error
    """
    try:
        # Validate required parameters
        subject = args.get("subject")
        email_account = args.get("email_account")
        
        if not subject:
            return {
                "error": "Missing required parameter: subject",
                "capability": "check_email_attachments"
            }
        
        if not email_account:
            return {
                "error": "Missing required parameter: email_account",
                "capability": "check_email_attachments"
            }
        
        # Get optional parameters
        search_unread_only = args.get("search_unread_only", True)
        
        # Check attachments
        processor = EmailProcessor()
        result = processor.check_email_attachments(
            subject=subject,
            email_account=email_account,
            search_unread_only=search_unread_only
        )
        
        if not result.get("email_found", False):
            return {
                "error": result.get("error", "Email not found"),
                "capability": "check_email_attachments"
            }
        
        return {
            "result": result,
            "capability": "check_email_attachments"
        }
        
    except Exception as e:
        logger.error(f"Error in check_email_attachments: {str(e)}")
        return {
            "error": str(e),
            "capability": "check_email_attachments"
        }


def check_specific_files(args: Dict[str, Any]) -> Dict[str, Any]:
    """
    Check if email has specific file patterns in attachments
    
    Args:
        args: Dictionary with capability arguments
        
    Returns:
        Dictionary with result or error
    """
    try:
        # Validate required parameters
        subject = args.get("subject")
        email_account = args.get("email_account")
        file_patterns = args.get("file_patterns")
        
        if not subject:
            return {
                "error": "Missing required parameter: subject",
                "capability": "check_specific_files"
            }
        
        if not email_account:
            return {
                "error": "Missing required parameter: email_account",
                "capability": "check_specific_files"
            }
        
        if not file_patterns:
            return {
                "error": "Missing required parameter: file_patterns",
                "capability": "check_specific_files"
            }
        
        if not isinstance(file_patterns, list):
            return {
                "error": "file_patterns must be a list of strings",
                "capability": "check_specific_files"
            }
        
        if len(file_patterns) == 0:
            return {
                "error": "file_patterns list cannot be empty",
                "capability": "check_specific_files"
            }
        
        # Get optional parameters
        search_unread_only = args.get("search_unread_only", True)
        
        # Check specific files
        processor = EmailProcessor()
        result = processor.check_specific_files(
            subject=subject,
            email_account=email_account,
            file_patterns=file_patterns,
            search_unread_only=search_unread_only
        )
        
        if not result.get("email_found", False):
            return {
                "error": result.get("error", "Email not found"),
                "capability": "check_specific_files"
            }
        
        return {
            "result": result,
            "capability": "check_specific_files"
        }
        
    except Exception as e:
        logger.error(f"Error in check_specific_files: {str(e)}")
        return {
            "error": str(e),
            "capability": "check_specific_files"
        }


def send_email_reply(args: Dict[str, Any]) -> Dict[str, Any]:
    """
    Send email reply
    
    Args:
        args: Dictionary with capability arguments
        
    Returns:
        Dictionary with result or error
    """
    try:
        # Validate required parameters
        to_email = args.get("to_email")
        subject = args.get("subject")
        body = args.get("body")
        
        if not to_email:
            return {
                "error": "Missing required parameter: to_email",
                "capability": "send_email_reply"
            }
        
        if not subject:
            return {
                "error": "Missing required parameter: subject",
                "capability": "send_email_reply"
            }
        
        if not body:
            return {
                "error": "Missing required parameter: body",
                "capability": "send_email_reply"
            }
        
        # Get optional parameters
        email_account = args.get("email_account")
        
        # Send email
        sender = EmailSender()
        result = sender.send_reply(
            to_email=to_email,
            subject=subject,
            body=body,
            email_account=email_account
        )
        
        if result.get("success", False):
            return {
                "result": result,
                "capability": "send_email_reply"
            }
        else:
            return {
                "error": result.get("error", "Failed to send email"),
                "capability": "send_email_reply"
            }
        
    except Exception as e:
        logger.error(f"Error in send_email_reply: {str(e)}")
        return {
            "error": str(e),
            "capability": "send_email_reply"
        }


def main():
    """Main entry point - reads JSON from stdin, outputs JSON to stdout"""
    try:
        # Read input from stdin
        input_data = json.load(sys.stdin)
        
        capability = input_data.get("capability")
        args = input_data.get("args", {})
        
        if not capability:
            print(json.dumps({
                "error": "Missing 'capability' in input",
                "capability": "unknown"
            }, indent=2))
            sys.exit(1)
        
        # Route to appropriate capability
        if capability == "find_and_extract_email":
            result = find_and_extract_email(args)
            print(json.dumps(result, indent=2))
        
        elif capability == "check_email_attachments":
            result = check_email_attachments(args)
            print(json.dumps(result, indent=2))
        
        elif capability == "check_specific_files":
            result = check_specific_files(args)
            print(json.dumps(result, indent=2))
        
        elif capability == "send_email_reply":
            result = send_email_reply(args)
            print(json.dumps(result, indent=2))
        
        else:
            print(json.dumps({
                "error": f"Unknown capability: {capability}",
                "capability": capability
            }, indent=2))
            sys.exit(1)
    
    except json.JSONDecodeError as e:
        print(json.dumps({
            "error": f"Invalid JSON input: {str(e)}",
            "capability": "unknown"
        }, indent=2))
        sys.exit(1)
    
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        print(json.dumps({
            "error": f"Error: {str(e)}",
            "capability": "unknown"
        }, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
