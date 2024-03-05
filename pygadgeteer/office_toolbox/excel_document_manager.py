import logging
import os
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
                # this could fail if the Excel is open previous to this call - then you can't change the visibility flag
                self.app.Visible = visible
            except Exception as e:
                logging.warning(
                    f"Exception : {e} when opening {self.filename =} can't change visibility"
                )

            try:
                self._document = self.app.Workbooks.Open(self.filename)
                self._new_document = False
            except pythoncom.com_error as error:
                logger.error(f"Error opening Excel document: {error}")
                self._document = None  # Ensure _document is None if open fails
        return self._document

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
            # the file exists
            os.remove(filename)
        if self._document:
            return self._document.SaveAs(filename, *argv, **kwargs)

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
