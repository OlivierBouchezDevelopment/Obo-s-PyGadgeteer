import logging
from typing import Optional

import pythoncom
from win32com.client import Dispatch, CDispatch

from .abstract_document_manager import AbstractDocumentManager

logger = logging.getLogger(__name__)


class ExcelDocumentManager(AbstractDocumentManager):
    """
    Manages Excel documents through COM for operations like opening, saving, and closing workbooks.

    Inherits from AbstractDocumentManager and implements methods for handling Excel documents.

    Attributes:
        filename (str): Path to the Excel workbook file being managed.
    """

    def __init__(self, filename: str):
        """
        Initializes the ExcelDocumentManager with a specific workbook file.

        Args:
            filename (str): Path to the Excel workbook file.
        """
        super().__init__(filename)
        self.app = Dispatch("Excel.Application", pythoncom.CoInitialize())

    def open_document(self, visible: bool = True) -> Optional[CDispatch]:
        """
        Opens the Excel document if not already opened. Sets the visibility of the Excel application
        based on the 'visible' argument.

        Args:
            visible (bool): Whether the Excel application should be visible to the user. Defaults to True.

        Returns:
            The opened Excel workbook.
        """
        if self._document is None:
            try:
                self.app.Visible = visible
                self._document = self.app.Workbooks.Open(self.filename)
                self._new_document = False
            except pythoncom.com_error as error:
                logger.error(f"Error opening Excel document: {error}")
                self._document = None  # Ensure _document is None if open fails
        return self._document

    def create_document(self, visible: bool = True) -> Optional[CDispatch]:
        """
        Creates a new Excel document. If a document is already opened, returns the active document object.

        Args:
            visible (bool): If True, the Excel application is made visible. Defaults to True.

        Returns:
            The newly created Excel document COM object.
        """
        if self._document is None:
            self.app.Visible = visible
            try:
                self._document = self.app.Workbooks.Add()
                self._new_document = True
            except pythoncom.com_error as error:
                logger.error(f"Error creating new document: {error}")
                self._document = None  # Ensure _document is None if open fails

        return self._document
