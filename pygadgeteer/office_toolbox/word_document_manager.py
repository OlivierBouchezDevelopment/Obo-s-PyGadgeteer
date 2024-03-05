import logging
from typing import Optional

import pythoncom
from win32com.client import Dispatch, CDispatch

from .abstract_document_manager import AbstractDocumentManager

logger = logging.getLogger(__name__)


class WordDocumentManager(AbstractDocumentManager):
    """
    Manages Word documents through COM for operations like opening, saving, and closing documents.

    Inherits from AbstractDocumentManager to provide specific functionality for managing Microsoft Word documents.

    Attributes:
        filename (str): Path to the document file being managed.
    """

    def __init__(self, filename: str):
        """
        Initializes the WordDocumentManager with a specific document file.

        Args:
            filename (str): Path to the Word document file.
        """
        super().__init__(filename)
        # Initialize the Word application COM object with automatic COM threading model initialization.
        self.app = Dispatch("Word.Application", pythoncom.CoInitialize())

    def open_document(self, visible: bool = True) -> Optional[CDispatch]:
        """
        Opens the specified Word document. If the document is already opened, returns the active document object.

        Args:
            visible (bool): If True, the Word application is made visible. Defaults to True.

        Returns:
            The opened Word document COM object.
        """
        if self._document is None:
            self.app.Visible = visible
            try:
                self._document = self.app.Documents.Open(self.filename)
                self._new_document = False
            except pythoncom.com_error as error:
                logger.error(f"Error opening document: {error}")
                self._document = None  # Ensure _document is None if open fails
        return self._document

    def create_document(self, visible: bool = True) -> Optional[CDispatch]:
        """
        Creates a new Word document. If a document is already opened, returns the active document object.

        Args:
            visible (bool): If True, the Word application is made visible. Defaults to True.

        Returns:
            The newly created Word document COM object.
        """
        if self._document is None:
            self.app.Visible = visible
            try:
                self._document = self.app.Documents.Add()
                self._new_document = True
            except pythoncom.com_error as error:
                logger.error(f"Error creating new document: {error}")
                self._document = None  # Ensure _document is None if open fails
        return self._document
