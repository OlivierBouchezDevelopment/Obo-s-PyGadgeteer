from abc import ABC, abstractmethod
import logging
from typing import Optional

import pythoncom
from win32com.client import CDispatch

logger = logging.getLogger(__name__)


class AbstractDocumentManager(ABC):
    """
    An abstract class for managing Office documents through COM for operations like opening,
    saving, and closing.

    Attributes:
        filename (str): Path to the document file being managed.
    """

    def __init__(self, filename: str):
        """
        Initializes the DocumentManager with a specific document file.

        Args:
            filename (str): Path to the document file.
        """
        self.filename = filename
        self.app = None
        self._document = None
        self._new_document = None

    @abstractmethod
    def open_document(self, visible: bool = True) -> Optional[CDispatch]:
        """
        Opens the document if not already opened.

        Args:
            visible (bool): Whether the document should be visible to the user. Defaults to True.
        """
        pass

    @abstractmethod
    def create_document(self, visible: bool = True) -> Optional[CDispatch]:
        """
        Creates a new document. If a document is already opened, returns the active document object.

        Args:
            visible (bool): Whether the document should be visible to the user. Defaults to True.

        Returns:
            The newly created document COM object.
        """

    @abstractmethod
    def save_as_document(self, filename: str, *argv, **kwargs):
        """
        Give access to the save as method for a file
        excel: self._document.SaveAs(filename, *argv, **kwargs)
        word: self._document.SaveAs2(filename, *argv, **kwargs)
        """
        self.filename = filename

    def save_document(self) -> None:
        """
        Saves the document. If the document is new, uses SaveAs2 to specify the filename;
        otherwise, uses Save to save changes.
        """
        try:
            if self._document:
                if self._new_document:
                    self.save_as_document(
                        self.filename
                    )  # or self._document.SaveAs2(self.filename)
                else:
                    self._document.Save()
        except pythoncom.com_error as error:
            logger.error(f"Error saving document: {error}")

    def close_document(self, save: bool = True) -> None:
        """
        Closes the document, optionally saving changes.

        Args:
            save (bool): Whether to save changes before closing. Defaults to True.
        """
        try:
            if self._document:
                if save:
                    self.save_document()
                self._document.Close()
                self._document = None
        except pythoncom.com_error as error:
            logger.error(f"Error closing document: {error}")

    def quit(self) -> None:
        """
        Quits the application, closing the document if open.
        """
        self.close_document(save=False)
        if self.app:
            self.app.Quit()
            self.app = None

    @property
    def document(self):
        """
        Ensures the document is opened and returns it.

        Returns:
            The opened document.
        """
        if not self._document:
            self.open_document()
        return self._document
