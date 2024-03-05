import os
from typing import Type

from .abstract_document_manager import AbstractDocumentManager
from .excel_document_manager import ExcelDocumentManager
from .word_document_manager import WordDocumentManager

# Type annotation for the document manager class dictionary
DOCUMENT_FACTORY: dict[str, Type[AbstractDocumentManager]] = {
    ".docx": WordDocumentManager,
    ".xlsx": ExcelDocumentManager,
}


def document_manager_factory(fullpath: str) -> AbstractDocumentManager:
    """
    Factory function to create an appropriate document manager instance based on the file extension.

    This function determines whether to create a WordDocumentManager or ExcelDocumentManager
    based on the file extension of the provided full path.

    Args:
        fullpath (str): The full path to the document file, including its name and extension.

    Returns:
        AbstractDocumentManager: An instance of a subclass of AbstractDocumentManager appropriate
        for the type of document specified by the fullpath argument.

    Raises:
        NotImplementedError: If a document manager for the specified file extension is not implemented.
    """
    _, extension = os.path.splitext(fullpath)
    document_manager_class = DOCUMENT_FACTORY.get(extension)

    if document_manager_class:
        return document_manager_class(fullpath)
    else:
        raise NotImplementedError(
            f"Office Document Manager for {extension} is not implemented."
        )
