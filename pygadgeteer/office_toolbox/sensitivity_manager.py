"""
This module manages Office Document document sensitivity labels using the Windows OLE/COM object libraries.
It provides functionality to set and retrieve sensitivity labels of Office Document documents and to create
a sensitivity label definition from a set of Office Document documents.
"""

from typing import Dict, Any
from win32com.client import Dispatch, CDispatch


class SensitivityLabelManager:
    """
    Manages sensitivity labels for an Office Document document using the COM object model.

    Attributes:
        document (CDispatch): The Office Document document object this manager operates on.
    """

    # shttps://learn.microsoft.com/fr-fr/office/vba/api/office.sensitivitylabel.createlabelinfo

    def __init__(self, document: CDispatch):
        """
        Initializes SensitivityLabelManager with an Office Document document.

        Args:
            document (CDispatch): The Office Document document object.
        """
        self.document = document

    @property
    def sensitivitylabel(self) -> CDispatch:
        """
        Retrieves the SensitivityLabel object from the document.

        Returns:
            CDispatch: The SensitivityLabel COM object associated with the document.
        """
        return self.document.SensitivityLabel

    def createlabelinfo(self) -> CDispatch:
        """
        Creates a new label info object for assigning a sensitivity label.

        Returns:
            CDispatch: A new LabelInfo COM object for sensitivity label assignment.
        """
        # https://learn.microsoft.com/fr-fr/office/vba/api/office.sensitivitylabel.createlabelinfo
        return self.sensitivitylabel.CreateLabelInfo()

    def getlabel(self) -> CDispatch:
        """
        Retrieves the current sensitivity label of the document.

        Returns:
            CDispatch: The current LabelInfo COM object representing the document's sensitivity label.
        """
        return self.sensitivitylabel.GetLabel()

    def setlabel(self, labelinfo: CDispatch) -> None:
        """
        Assigns a sensitivity label to the document.

        Args:
            labelinfo (CDispatch): The LabelInfo object containing label details.
        """
        self.sensitivitylabel.SetLabel(labelinfo, labelinfo)


class LabelInfoManager:
    """
    Manages and inspects LabelInfo objects for sensitivity labels.
    """

    # https://learn.microsoft.com/fr-fr/office/vba/api/office.labelinfo
    def __init__(self, labelinfo: CDispatch):
        """
        Initializes LabelInfoManager with a LabelInfo COM object.

        Args:
            labelinfo (CDispatch): The LabelInfo COM object.
        """
        self.labelinfo = labelinfo

    def dump_info(self) -> Dict[str, Any]:
        """
        Extracts and returns all available information from the LabelInfo object.

        Returns:
            Dict[str, Any]: A dictionary containing attributes and their values from the LabelInfo object.
        """
        dump = {}
        for attr in dir(self.labelinfo):
            if attr.startswith("_"):  # Skip private or protected attributes
                continue
            value = getattr(self.labelinfo, attr)
            if isinstance(value, (str, int, bool, float)):  # Check for basic data types
                dump[attr] = value
        return dump
