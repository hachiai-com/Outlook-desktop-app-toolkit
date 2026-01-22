"""
Automated email reply functionality using Outlook COM API
"""
import logging
from typing import Optional, Dict, Any
from outlook_connector import OutlookConnector
from config import ToolkitConfig


class EmailSender:
    """Handles sending automated reply emails via Outlook"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.config = ToolkitConfig()
        self.connector = OutlookConnector()
    
    def send_reply(
        self,
        to_email: str,
        subject: str,
        body: str,
        email_account: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Send email reply via Outlook
        
        Args:
            to_email: Recipient email address
            subject: Email subject
            body: Email body
            email_account: Optional account to send from (uses default if not specified)
            
        Returns:
            Dictionary with send result
        """
        try:
            self.connector.initialize_com()
            outlook = self.connector.get_outlook_application()
            
            # Create mail item (0 = olMailItem)
            mail = outlook.CreateItem(0)
            
            # Set email properties
            mail.To = to_email
            mail.Subject = subject
            mail.Body = body
            
            # If email_account is specified, set the send account
            if email_account:
                # Get accounts collection
                namespace = outlook.GetNamespace("MAPI")
                accounts = namespace.Accounts
                
                # Find the account
                account_found = False
                for account in accounts:
                    if account.DisplayName == email_account or account.SmtpAddress == email_account:
                        mail.SendUsingAccount = account
                        account_found = True
                        self.logger.info(f"Using account: {account.DisplayName}")
                        break
                
                if not account_found:
                    self.logger.warning(f"Account {email_account} not found, using default account")
            
            # Send email
            mail.Send()
            
            self.logger.info(f"Successfully sent email to {to_email} with subject: {subject}")
            
            return {
                "success": True,
                "to": to_email,
                "subject": subject,
                "message": "Email sent successfully"
            }
            
        except Exception as e:
            error_msg = f"Error sending email: {str(e)}"
            self.logger.error(error_msg)
            return {
                "success": False,
                "error": error_msg
            }
        finally:
            self.connector.uninitialize_com()
    
    def send_attachment_request_reply(
        self,
        original_email: object,
        reply_message: Optional[str] = None,
        email_account: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Send automated reply requesting attachments
        
        Args:
            original_email: Original email object to reply to
            reply_message: Custom reply message (uses default template if not provided)
            email_account: Optional account to send from
            
        Returns:
            Dictionary with send result
        """
        try:
            # Get sender email from original email
            sender_email = getattr(original_email, 'SenderEmailAddress', None)
            if not sender_email:
                # Try alternative property
                sender_email = getattr(original_email, 'Sender', None)
                if hasattr(sender_email, 'Address'):
                    sender_email = sender_email.Address
            
            if not sender_email:
                return {
                    "success": False,
                    "error": "Could not determine sender email address"
                }
            
            # Get subject
            original_subject = getattr(original_email, 'Subject', 'Your email')
            
            # Create reply subject
            reply_subject = f"Re: {original_subject}"
            
            # Create reply body
            if reply_message is None:
                reply_message = self.config.DEFAULT_REPLY_MESSAGE.format(
                    subject=original_subject
                )
            
            # Send reply
            return self.send_reply(
                to_email=sender_email,
                subject=reply_subject,
                body=reply_message,
                email_account=email_account
            )
            
        except Exception as e:
            error_msg = f"Error sending attachment request reply: {str(e)}"
            self.logger.error(error_msg)
            return {
                "success": False,
                "error": error_msg
            }
