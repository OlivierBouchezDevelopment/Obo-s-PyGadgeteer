import json
from glob import glob
import os
from typing import Dict, Any, Type

from .abstract_document_manager import AbstractDocumentManager
from .sensitivity_manager import SensitivityLabelManager, LabelInfoManager
from .document_manager_factory import document_manager_factory

DEFAULT_SENSITIVITY_LABELS_DEFINITION = (
    "sensitivity_model/sensitivity_labels_definition.json"
)
DEFAULT_SENSITIVITY_TEMPLATES = "sensitivity_model"


def set_sensitivity_label_to_document(
    document_manager: AbstractDocumentManager,
    sensitivity_label: str,
    sensitivity_configuration_file: str = DEFAULT_SENSITIVITY_LABELS_DEFINITION,
) -> None:
    """
    Sets a sensitivity label to a document managed by the provided document manager instance,
    based on the predefined labels configuration.

    This function reads the sensitivity labels configuration file, creates a new label info object,
    sets the label properties based on the specified sensitivity label, and assigns it to the document.

    Args:
        document_manager (AbstractDocumentManager): The document manager responsible for the document to label.
        sensitivity_label (str): The key representing the sensitivity label to apply, as defined in the sensitivity labels configuration file.
        sensitivity_configuration_file (str): Path to the JSON file containing sensitivity labels configuration. Defaults to DEFAULT_SENSITIVITY_LABELS_DEFINITION.

    Raises:
        FileNotFoundError: If the specified sensitivity configuration file does not exist.
        KeyError: If the specified sensitivity label is not found in the configuration file.
    """
    with open(sensitivity_configuration_file, "r") as fh_in:
        sensitivity_labels = json.load(fh_in)

    if not document_manager.document:
        return
    sensitivity_label_manager = SensitivityLabelManager(document_manager.document)
    new_label_info = sensitivity_label_manager.createlabelinfo()
    new_label_info.AssignmentMethod = 2  # Manual assignment
    new_label_info.Justification = (
        f"Automated assignment based on configuration.({sensitivity_label})"
    )
    new_label_info.LabelId = sensitivity_labels[sensitivity_label]["LabelId"]
    new_label_info.LabelName = sensitivity_labels[sensitivity_label]["LabelName"]
    sensitivity_label_manager.setlabel(new_label_info)


def set_sensitivity_label_to_file(
    absolute_path_to_filename: str,
    sensitivity_label: str,
    sensitivity_configuration_file: str = DEFAULT_SENSITIVITY_LABELS_DEFINITION,
) -> None:
    """
    Sets the sensitivity label for a document file located at the specified path, using the provided sensitivity label.

    This function creates a document manager for the specified file, applies the specified sensitivity label using the configuration file, and then saves and closes the document. It supports documents that can be managed by the implemented document managers (e.g., Word, Excel).

    Args:
        absolute_path_to_filename (str): The full path to the document file. The file type should be supported by the available document managers.
        sensitivity_label (str): The sensitivity label to apply to the document. This should match a key in the sensitivity labels configuration file.
        sensitivity_configuration_file (str, optional): Path to the sensitivity labels configuration file. This file contains the mapping of sensitivity label keys to their respective label IDs and names. Defaults to DEFAULT_SENSITIVITY_LABELS_DEFINITION.

    Raises:
        NotImplementedError: If the document manager for the specified file type is not implemented.
        FileNotFoundError: If the specified sensitivity configuration file does not exist.
        KeyError: If the specified sensitivity label is not found in the configuration file.
    """
    document_manager = document_manager_factory(absolute_path_to_filename)
    set_sensitivity_label_to_document(
        document_manager,
        sensitivity_label,
        sensitivity_configuration_file,
    )
    document_manager.close_document(save=True)
    document_manager.quit()


def create_sensitivity_label_definition(
    extract_from: str = DEFAULT_SENSITIVITY_TEMPLATES,
    sensitivity_configuration_file: str = DEFAULT_SENSITIVITY_LABELS_DEFINITION,
) -> None:
    """
    Creates a sensitivity label definition file from a set of Office documents.

    Args:
        extract_from: Directory containing Office documents to extract labels from.
        sensitivity_configuration_file: Path to save the generated configuration file.

    Note:
        You have to create in extract_from - ex: sensitivity_model as series of files with the different sensitivity level you want to capture.
        Ex:
        create public.xlxs (with sensitivity level public)
        create internaluseonly.xlsx (with sensitivity level internal use only).

        Each company have it's own sensitivity level and those have specific ids.

    """
    sensitivity_label_definition: Dict[str, Any] = {}

    for filename in glob(os.path.join(extract_from, "*")):
        label, extension = os.path.splitext(os.path.basename(filename))
        if extension not in [".xlsx", ".docx"]:
            continue

        document_manager = document_manager_factory(os.path.abspath(filename))
        if not document_manager.document:
            continue

        sensitivity_label_manager = SensitivityLabelManager(document_manager.document)
        current_label_info = sensitivity_label_manager.getlabel()
        label_info_manager = LabelInfoManager(current_label_info)

        sensitivity_label_definition[label] = {
            "LabelId": label_info_manager.labelinfo.LabelId,
            "LabelName": label_info_manager.labelinfo.LabelName,
            "IsEnabled": label_info_manager.labelinfo.IsEnabled,
            "SetDate": label_info_manager.labelinfo.SetDate,
            "AssignmentMethod": label_info_manager.labelinfo.AssignmentMethod,
            "SiteId": label_info_manager.labelinfo.SiteId,
            "ActionId": label_info_manager.labelinfo.ActionId,
            "ContentBits": label_info_manager.labelinfo.ContentBits,
        }

        document_manager.close_document(save=False)

    with open(sensitivity_configuration_file, "w", encoding="utf8") as fh_out:
        json.dump(sensitivity_label_definition, fh_out, indent=4)
