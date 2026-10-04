"""
This module houses interactions with Outlook (Classic) via `pywin32`.

Here is a note of the COM object structure for Outlook.
Hierarchy:
Outlook.Application
└── Namespace (MAPI)
    ├── Store 1 (email address 1)
    │   └── Root Folder
    │       ├── Inbox
    │       │   ├── Mail Item 1
    │       │   │   └── Attachments
    │       │   ├── Mail Item 2
    │       │   └── ...
    │       ├── Sent Items
    │       ├── Drafts
    │       └── Custom Folder
    │           └── SubFolder
    │               └── Mail Items...
    │
    ├── Store 2 (email address 2)
    │   └── Root Folder
    │       ├── Inbox
    │       ├── ...
    │
    └── Store N
        └── ...
"""


import os
import logging
import win32com.client as win32
from pathlib import Path


# Configure logging
from .log import default_log_config
default_log_config()


def get_app():
    return win32.Dispatch("Outlook.Application")

def get_mapi_namespace():
    olk_app = get_app()
     # MAPI means Messaging API
     # (a low-level API for accessing messaging systems, including email, calendar, contacts, etc.)
    return olk_app.GetNamespace("MAPI")

def get_all_stores():
    """
    Returns a dict of all backend stores configured in Outlook.
    Store here means the stored data and settings, it's a legacy MAPI concept.
    Returns:
        dict: {display_name: store_object, ...}
              Example: {'user@example.com': store_obj, 'archive@example.com': store_obj}
    """

    olk_namespace = get_mapi_namespace()
    stores = {}
    
    for store in olk_namespace.Stores:
        display_name = store.DisplayName
        stores[display_name] = store
        logging.info(f"Found Store: {display_name}")
    
    return stores

def get_store(store_name: str = None):
    """
    Returns a backend store configured in Outlook.
    Store here means the stored data and settings, it's a legacy MAPI concept.
    """
    olk_namespace = get_mapi_namespace()

    if store_name is None:
        # Get the default store (the one with the default email account)
        default_store = olk_namespace.DefaultStore
        return default_store

    for store in olk_namespace.Stores:
        if store.DisplayName == store_name:
            logging.info(f"Found Store: {store_name}")
            return store
        
    raise ValueError(f"Store '{store_name}' not found in Outlook.")

def get_root_folder(store_name: str = None):
    if store_name is None:
        # Get the default store (the one with the default email account)
        namespace = get_mapi_namespace()
        default_store = namespace.DefaultStore
        logging.info(f"Using default Store: {default_store.DisplayName}")
        return default_store.GetRootFolder()
    else:
        store = get_store(store_name)
        return store.GetRootFolder()

def get_inbox(store_name: str = None):
    root_folder = get_root_folder(store_name)
    # 6 corresponds to the Inbox folder
    return root_folder.Folders.Item(6)

def get_folder(*, path: str, store_name: str = None):
    """
    Returns a mailbox folder by name. If mailbox_path contains path separators ('/'), it will be treated as a path and searched recursively.

    Args:
        mailbox_path (str):
        store_name (str):
    Returns:
        mailbox_folder:
    """
    root_folder = get_root_folder(store_name)

    if "/" in path:
        path_parts = path.split("/")
        current_folder = root_folder
        for part in path_parts:
            try:
                current_folder = current_folder.Folders.Item(part)
            except ValueError:
                raise ValueError(f"Mailbox folder '{part}' not found in path '{path}'.")
        logging.info(f"Found Mailbox Folder: {path}")
        return current_folder
    else:
        try:
            mailbox_folder = root_folder.Folders.Item(path)
            logging.info(f"Found Mailbox Folder: {path}")
            return mailbox_folder
        except ValueError:
            raise ValueError(f"Mailbox '{path}' not found in Outlook.")

def save_attachments(mail_item, save_path: os.PathLike | str | Path):
    save_path = Path(save_path).resolve()
    save_path.mkdir(parents=True, exist_ok=True)

    attachments = mail_item.Attachments
    if attachments.Count == 0:
        logging.warning(f"No attachments found for email '{getattr(mail_item, 'Subject', '<no subject>')}'.")
        return

    for attachment in attachments:
        attachment_path = save_path / attachment.FileName
        attachment.SaveAsFile(str(attachment_path))
        logging.info(f"Saved attachment: {attachment_path}")


# Search emails
# 1. If subject is known, filter for exact subject match.
# 2. If search by keywords, filter for emails that contain all keywords in the subject or body.
def search_emails(mailbox_folder, subject: str = None, keywords: list = None):
    if subject is not None:
        # Filter for exact subject match
        filtered_items = mailbox_folder.Items.Restrict(f"[Subject] = '{subject}'")
    else:
        filtered_items = mailbox_folder.Items

    if keywords is not None:
        # Filter for emails containing all keywords
        for keyword in keywords:
            filtered_items = filtered_items.Restrict(f"[[Subject] LIKE '{keyword}' OR [Body] LIKE '{keyword}']")
    
    return filtered_items