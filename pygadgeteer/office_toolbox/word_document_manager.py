import logging
import os
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

    def save_as_document(self, filename: str, *argv, **kwargs):
        """
        Saves the current document under a new name, handling the differences between Excel and Word save methods.

        This method abstracts the discrepancies between Excel's `SaveAs` and Word's `SaveAs2` methods, allowing for
        a unified interface to save documents in both applications. If the specified file already exists, it is
        overwritten without prompting. This method adjusts for Word's requirement by using `SaveAs2`, ensuring
        compatibility and enabling the passing of additional arguments and keyword arguments specific to each application's
        save method.

        Args:
            filename (str): The path, including the name of the file, where the document will be saved. If a file with
                            the same name exists, it will be overwritten.
            *argv: Variable length argument list passed directly to the Excel `SaveAs` or Word `SaveAs2` method,
                allowing for application-specific save options (e.g., file format or password protection).
            **kwargs: Arbitrary keyword arguments passed directly to the Excel `SaveAs` or Word `SaveAs2` method,
                    supporting a wide range of saving options specific to each application.

        Returns:
            The result of the save operation, typically `None` unless the underlying method provides a return value.

        Raises:
            Various exceptions can be raised depending on the application and the arguments provided. Common issues
            include COM errors due to invalid file paths or permissions, and application-specific errors related to
            save options.

        Note:
            This method does not distinguish between Excel and Word internally; it uses `SaveAs2` universally,
            which is intended for Word but may work with Excel in some contexts. Ensure compatibility with your
            specific use case and Office application version.
        """
        self.filename = filename
        if os.path.exists(filename):
            # the file exists.. Is there an option to force writing over an empty doc with SaveAs2 ?
            os.remove(filename)

        if self._document:
            return self._document.SaveAs2(filename, *argv, **kwargs)

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
