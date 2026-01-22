"""
Configuration management for Outlook Desktop Toolkit
"""
import os
from pathlib import Path
from typing import Optional
from datetime import datetime
import re


class ToolkitConfig:
    """Configuration settings for the Outlook Desktop Toolkit"""
    
    # Default settings
    DEFAULT_OUTPUT_BASE_PATH = os.getcwd()
    DEFAULT_REPLY_MESSAGE = "Please provide the required attachments for: {subject}"
    DEFAULT_SEARCH_UNREAD_ONLY = True
    DEFAULT_SEND_REPLY_IF_NO_ATTACHMENTS = False
    
    def __init__(self):
        self.output_base_path = self.DEFAULT_OUTPUT_BASE_PATH
        self.default_reply_message = self.DEFAULT_REPLY_MESSAGE
        self.search_unread_only = self.DEFAULT_SEARCH_UNREAD_ONLY
        self.send_reply_if_no_attachments = self.DEFAULT_SEND_REPLY_IF_NO_ATTACHMENTS
    
    @staticmethod
    def clean_filename(filename: str, max_length: int = 200) -> str:
        """
        Remove invalid characters from filename and limit length
        
        Args:
            filename: Original filename
            max_length: Maximum length for filename
            
        Returns:
            Cleaned filename safe for filesystem
        """
        # Remove invalid characters for Windows filesystem
        invalid_chars = r'[:\\/?*"<>|]'
        cleaned = re.sub(invalid_chars, '', filename)
        # Replace spaces with underscores for better compatibility
        cleaned = cleaned.replace(' ', '_')
        # Limit length and trim whitespace
        cleaned = cleaned.strip()[:max_length]
        return cleaned
    
    @staticmethod
    def generate_timestamp() -> str:
        """
        Generate timestamp string for folder naming
        
        Returns:
            Timestamp string in format YYYY-MM-DD_HH-MM-SS
        """
        return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    @staticmethod
    def create_email_folder_name(subject: str, timestamp: Optional[str] = None) -> str:
        """
        Create folder name for email extraction
        
        Args:
            subject: Email subject
            timestamp: Optional timestamp (generated if not provided)
            
        Returns:
            Folder name combining cleaned subject and timestamp
        """
        if timestamp is None:
            timestamp = ToolkitConfig.generate_timestamp()
        
        cleaned_subject = ToolkitConfig.clean_filename(subject, max_length=100)
        return f"{cleaned_subject}_{timestamp}"
