"""
Email search, extraction, and attachment handling
"""
import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
import logging
from datetime import datetime
from outlook_connector import OutlookConnector
from config import ToolkitConfig


class EmailProcessor:
    """Handles email search, content extraction, and attachment downloading"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.config = ToolkitConfig()
        self.connector = OutlookConnector()
    
    def find_email_by_subject(
        self,
        subject: str,
        email_account: str,
        search_unread_only: bool = True
    ) -> Optional[object]:
        """
        Find most recent email matching subject (case-insensitive contains)
        
        Args:
            subject: Email subject to search for
            email_account: Outlook email account ID
            search_unread_only: If True, search only unread emails
            
        Returns:
            Email object if found, None otherwise
        """
        try:
            self.connector.initialize_com()
            inbox, _ = self.connector.get_inbox(email_account)
            
            # Get items (unread or all)
            if search_unread_only:
                items = inbox.Items.Restrict("[Unread] = True")
                self.logger.info(f"Searching unread emails for subject containing: {subject}")
            else:
                items = inbox.Items
                self.logger.info(f"Searching all emails for subject containing: {subject}")
            
            # Sort by ReceivedTime descending (most recent first)
            items.Sort("[ReceivedTime]", True)
            
            # Search for matching subject (case-insensitive contains)
            subject_lower = subject.lower()
            email = items.GetFirst()
            
            while email:
                try:
                    if email.Class == 43:  # MailItem
                        email_subject = getattr(email, 'Subject', '')
                        if subject_lower in email_subject.lower():
                            self.logger.info(f"Found matching email: {email_subject}")
                            return email
                except Exception as e:
                    self.logger.warning(f"Error checking email: {str(e)}")
                    continue
                
                email = items.GetNext()
            
            self.logger.warning(f"No email found with subject containing: {subject}")
            return None
            
        except Exception as e:
            self.logger.error(f"Error searching for email: {str(e)}")
            raise
        finally:
            self.connector.uninitialize_com()
    
    def extract_email_content(self, email: object) -> Dict[str, str]:
        """
        Extract email content and metadata
        
        Args:
            email: Outlook email object
            
        Returns:
            Dictionary with email metadata and body
        """
        try:
            return {
                "subject": getattr(email, 'Subject', 'No Subject'),
                "sender_name": getattr(email, 'SenderName', 'Unknown'),
                "sender_email": getattr(email, 'SenderEmailAddress', 'Unknown'),
                "to": getattr(email, 'To', 'Unknown'),
                "cc": getattr(email, 'CC', ''),
                "sent_on": str(getattr(email, 'SentOn', 'Unknown')),
                "received_time": str(getattr(email, 'ReceivedTime', 'Unknown')),
                "body": getattr(email, 'Body', ''),
                "entry_id": getattr(email, 'EntryID', ''),
            }
        except Exception as e:
            self.logger.error(f"Error extracting email content: {str(e)}")
            raise
    
    def save_email_content(
        self,
        email_data: Dict[str, str],
        output_path: Path
    ) -> str:
        """
        Save email content to text file
        
        Args:
            email_data: Dictionary with email metadata and body
            output_path: Path to save the text file
            
        Returns:
            Full path to saved file
        """
        try:
            content_file = output_path / "email_content.txt"
            
            with open(content_file, 'w', encoding='utf-8') as f:
                f.write(f"Subject: {email_data['subject']}\n")
                f.write(f"From: {email_data['sender_name']} <{email_data['sender_email']}>\n")
                f.write(f"To: {email_data['to']}\n")
                if email_data.get('cc'):
                    f.write(f"CC: {email_data['cc']}\n")
                f.write(f"Sent: {email_data['sent_on']}\n")
                f.write(f"Received: {email_data['received_time']}\n")
                f.write("-" * 80 + "\n")
                f.write("Body:\n")
                f.write(email_data['body'])
            
            self.logger.info(f"Saved email content to: {content_file}")
            return str(content_file)
            
        except Exception as e:
            self.logger.error(f"Error saving email content: {str(e)}")
            raise
    
    def download_attachments(
        self,
        email: object,
        attachments_folder: Path
    ) -> List[Dict[str, Any]]:
        """
        Download all attachments from email
        
        Args:
            email: Outlook email object
            attachments_folder: Path to save attachments
            
        Returns:
            List of dictionaries with attachment information
        """
        attachments_info = []
        
        try:
            attachment_count = email.Attachments.Count
            
            if attachment_count == 0:
                self.logger.info("Email has no attachments")
                return attachments_info
            
            # Create attachments folder if it doesn't exist
            attachments_folder.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"Downloading {attachment_count} attachment(s)")
            
            for i in range(1, attachment_count + 1):  # Outlook is 1-indexed
                try:
                    attachment = email.Attachments.Item(i)
                    filename = attachment.FileName
                    
                    # Clean filename for filesystem safety
                    cleaned_filename = self.config.clean_filename(filename)
                    file_path = attachments_folder / cleaned_filename
                    
                    # Save attachment
                    attachment.SaveAsFile(str(file_path))
                    
                    # Get file size
                    file_size = file_path.stat().st_size if file_path.exists() else 0
                    
                    attachments_info.append({
                        "filename": filename,
                        "cleaned_filename": cleaned_filename,
                        "path": str(file_path),
                        "size_bytes": file_size
                    })
                    
                    self.logger.info(f"Downloaded attachment: {filename} -> {file_path}")
                    
                except Exception as e:
                    self.logger.error(f"Error downloading attachment {i}: {str(e)}")
                    continue
            
            return attachments_info
            
        except Exception as e:
            self.logger.error(f"Error processing attachments: {str(e)}")
            raise
    
    def process_email(
        self,
        subject: str,
        email_account: str,
        output_base_path: Optional[str] = None,
        search_unread_only: bool = True
    ) -> Dict[str, Any]:
        """
        Main method to find, extract, and download email with attachments
        
        Args:
            subject: Email subject to search for
            email_account: Outlook email account ID
            output_base_path: Base directory for output (default: current directory)
            search_unread_only: If True, search only unread emails
            
        Returns:
            Dictionary with processing results
        """
        try:
            # Set output path
            if output_base_path is None:
                output_base_path = self.config.DEFAULT_OUTPUT_BASE_PATH
            
            # Find email
            email = self.find_email_by_subject(subject, email_account, search_unread_only)
            
            if email is None:
                return {
                    "email_found": False,
                    "error": f"No email found with subject containing: {subject}"
                }
            
            # Extract email content
            email_data = self.extract_email_content(email)
            
            # Create folder structure
            email_folder_name = self.config.create_email_folder_name(
                email_data['subject']
            )
            base_extraction_path = Path(output_base_path) / "email_extractions"
            email_folder = base_extraction_path / email_folder_name
            email_folder.mkdir(parents=True, exist_ok=True)
            
            # Save email content
            email_content_file = self.save_email_content(email_data, email_folder)
            
            # Download attachments
            attachments_folder = email_folder / "attachments"
            attachments_info = self.download_attachments(email, attachments_folder)
            
            # Build result
            result = {
                "email_found": True,
                "email_subject": email_data['subject'],
                "email_sender": email_data['sender_name'],
                "email_sender_address": email_data['sender_email'],
                "email_sent_time": email_data['sent_on'],
                "email_received_time": email_data['received_time'],
                "has_attachments": len(attachments_info) > 0,
                "attachment_count": len(attachments_info),
                "output_folder": str(email_folder),
                "email_content_file": email_content_file,
                "attachments_folder": str(attachments_folder),
                "attachments": attachments_info,
                "reply_sent": False  # Will be set by email_sender if reply is sent
            }
            
            self.logger.info(f"Successfully processed email: {email_data['subject']}")
            return result
            
        except Exception as e:
            error_msg = f"Error processing email: {str(e)}"
            self.logger.error(error_msg)
            return {
                "email_found": False,
                "error": error_msg
            }
    
    def check_email_attachments(
        self,
        subject: str,
        email_account: str,
        search_unread_only: bool = True
    ) -> Dict[str, Any]:
        """
        Check if email has any attachments
        
        Args:
            subject: Email subject to search for
            email_account: Outlook email account ID
            search_unread_only: If True, search only unread emails
            
        Returns:
            Dictionary with attachment check results
        """
        try:
            # Find email
            email = self.find_email_by_subject(subject, email_account, search_unread_only)
            
            if email is None:
                return {
                    "email_found": False,
                    "error": f"No email found with subject containing: {subject}"
                }
            
            # Extract email metadata
            email_data = self.extract_email_content(email)
            
            # Check attachments
            attachment_count = email.Attachments.Count
            has_attachments = attachment_count > 0
            
            # Get attachment list
            attachments_list = []
            if has_attachments:
                for i in range(1, attachment_count + 1):  # Outlook is 1-indexed
                    try:
                        attachment = email.Attachments.Item(i)
                        attachments_list.append({
                            "filename": attachment.FileName,
                            "size_bytes": getattr(attachment, 'Size', 0)
                        })
                    except Exception as e:
                        self.logger.warning(f"Error getting attachment {i} info: {str(e)}")
                        continue
            
            return {
                "email_found": True,
                "email_subject": email_data['subject'],
                "email_sender": email_data['sender_name'],
                "email_sent_time": email_data['sent_on'],
                "has_attachments": has_attachments,
                "attachment_count": attachment_count,
                "attachments": attachments_list
            }
            
        except Exception as e:
            error_msg = f"Error checking email attachments: {str(e)}"
            self.logger.error(error_msg)
            return {
                "email_found": False,
                "error": error_msg
            }
    
    def check_specific_files(
        self,
        subject: str,
        email_account: str,
        file_patterns: List[str],
        search_unread_only: bool = True
    ) -> Dict[str, Any]:
        """
        Check if email has specific file patterns in attachments
        
        Args:
            subject: Email subject to search for
            email_account: Outlook email account ID
            file_patterns: List of file name patterns to search for (case-insensitive)
            search_unread_only: If True, search only unread emails
            
        Returns:
            Dictionary with file pattern check results
        """
        try:
            # Find email
            email = self.find_email_by_subject(subject, email_account, search_unread_only)
            
            if email is None:
                return {
                    "email_found": False,
                    "error": f"No email found with subject containing: {subject}"
                }
            
            # Extract email metadata
            email_data = self.extract_email_content(email)
            
            # Check attachments
            attachment_count = email.Attachments.Count
            has_attachments = attachment_count > 0
            
            # Get all attachment filenames
            attachment_filenames = []
            if has_attachments:
                for i in range(1, attachment_count + 1):  # Outlook is 1-indexed
                    try:
                        attachment = email.Attachments.Item(i)
                        attachment_filenames.append(attachment.FileName)
                    except Exception as e:
                        self.logger.warning(f"Error getting attachment {i} info: {str(e)}")
                        continue
            
            # Check which patterns are found
            found_patterns = []
            missing_patterns = []
            pattern_details = {}
            
            for pattern in file_patterns:
                pattern_lower = pattern.lower()
                found = False
                matching_files = []
                
                # Check if pattern matches any attachment filename
                for filename in attachment_filenames:
                    filename_lower = filename.lower()
                    if pattern_lower in filename_lower:
                        found = True
                        matching_files.append(filename)
                
                if found:
                    found_patterns.append(pattern)
                    pattern_details[pattern] = {
                        "found": True,
                        "matching_files": matching_files
                    }
                else:
                    missing_patterns.append(pattern)
                    pattern_details[pattern] = {
                        "found": False,
                        "matching_files": []
                    }
            
            return {
                "email_found": True,
                "email_subject": email_data['subject'],
                "email_sender": email_data['sender_name'],
                "email_sent_time": email_data['sent_on'],
                "has_attachments": has_attachments,
                "attachment_count": attachment_count,
                "all_attachments": attachment_filenames,
                "patterns_searched": file_patterns,
                "found_patterns": found_patterns,
                "missing_patterns": missing_patterns,
                "pattern_details": pattern_details,
                "all_patterns_found": len(missing_patterns) == 0
            }
            
        except Exception as e:
            error_msg = f"Error checking specific files: {str(e)}"
            self.logger.error(error_msg)
            return {
                "email_found": False,
                "error": error_msg
            }
