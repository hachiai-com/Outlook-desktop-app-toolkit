"""
Outlook COM API connection wrapper for thread-safe access
"""
import pythoncom
import win32com.client
from typing import Tuple, Optional
import logging


class OutlookConnector:
    """Manages connection to Outlook desktop application via COM API"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self._outlook = None
        self._namespace = None
        self._com_initialized = False
    
    def initialize_com(self):
        """Initialize COM for current thread"""
        if not self._com_initialized:
            try:
                pythoncom.CoInitialize()
                self._com_initialized = True
                self.logger.debug("COM initialized for thread")
            except Exception as e:
                self.logger.error(f"Failed to initialize COM: {str(e)}")
                raise
    
    def uninitialize_com(self):
        """Uninitialize COM for current thread"""
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
                self._com_initialized = False
                self.logger.debug("COM uninitialized for thread")
            except Exception as e:
                self.logger.error(f"Failed to uninitialize COM: {str(e)}")
    
    def connect(self) -> Tuple[object, object]:
        """
        Connect to Outlook application and return outlook and namespace objects
        
        Returns:
            Tuple of (outlook, namespace) objects
            
        Raises:
            Exception: If connection fails
        """
        try:
            if self._outlook is None:
                self._outlook = win32com.client.Dispatch("Outlook.Application")
                self.logger.debug("Connected to Outlook application")
            
            if self._namespace is None:
                self._namespace = self._outlook.GetNamespace("MAPI")
                self.logger.debug("Accessed MAPI namespace")
            
            return self._outlook, self._namespace
            
        except Exception as e:
            error_msg = f"Failed to connect to Outlook: {str(e)}"
            self.logger.error(error_msg)
            raise Exception(error_msg)
    
    def get_inbox(self, email_account: str):
        """
        Get inbox folder for specified email account
        
        Args:
            email_account: Email account ID (e.g., "test@gmail.com")
            
        Returns:
            Tuple of (inbox, outlook) objects
            
        Raises:
            Exception: If inbox access fails
        """
        try:
            if not email_account:
                raise Exception("Email account ID is required")
            
            outlook, namespace = self.connect()
            
            # Access the specific account's folder
            account_folder = namespace.Folders(email_account)
            if account_folder is None:
                raise Exception(f"Could not find folder for account: {email_account}")
            
            # Get Inbox folder
            inbox = account_folder.Folders("Inbox")
            if inbox is None:
                raise Exception(f"Could not find Inbox folder for account: {email_account}")
            
            self.logger.debug(f"Successfully accessed inbox for account: {email_account}")
            return inbox, outlook
            
        except Exception as e:
            error_msg = f"Failed to access inbox for account {email_account}: {str(e)}"
            self.logger.error(error_msg)
            raise Exception(error_msg)
    
    def get_outlook_application(self):
        """
        Get Outlook application object
        
        Returns:
            Outlook application object
        """
        outlook, _ = self.connect()
        return outlook
    
    def __enter__(self):
        """Context manager entry"""
        self.initialize_com()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit"""
        self.uninitialize_com()
