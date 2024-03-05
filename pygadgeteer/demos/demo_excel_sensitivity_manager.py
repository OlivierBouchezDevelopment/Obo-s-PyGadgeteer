# To run this demo, you need to have Python installed on your computer.
# First, install the pywin32 library by running the following command in your terminal or command prompt:
# pip install pywin32

import os
import json
import shutil

# Import the necessary modules from the office_toolbox package.
# Make sure the office_toolbox package is correctly installed or located in your project directory.
from office_toolbox.document_manager_factory import document_manager_factory
from office_toolbox.sensitivity_manager import SensitivityLabelManager, LabelInfoManager
from office_toolbox.set_sensitivity_label import (
    create_sensitivity_label_definition,
    set_sensitivity_label_to_file,
    DEFAULT_SENSITIVITY_TEMPLATES,
    DEFAULT_SENSITIVITY_LABELS_DEFINITION,
)


def demo_show_label_info():
    """
    This demo function iterates through a list of filenames and their corresponding sensitivity labels.
    It opens each document, retrieves its current sensitivity label, and prints the label information.
    """
    document_manager = None
    for filename, sensitivity in [
        ("output/dummy_internal_use_only.xlsx", "InternalUseOnly"),
        ("output/dummy_commercial_in_confidence.xlsx", "CommercialInConfidence"),
        ("output/dummy_public.xlsx", "Public"),
        ("output/dummy_restricted_confidential.xlsx", "RestrictedConfidential"),
        ("output/dummy_restricted_sensitive.xlsx", "RestrictedSensitive"),
    ]:
        print(f"{filename = },{sensitivity = }")
        fullpath = os.path.join(os.getcwd(), filename)
        document_manager = document_manager_factory(fullpath)
        if document_manager.document is None:
            continue
        sensitivity_manager = SensitivityLabelManager(document_manager.document)
        label_info = LabelInfoManager(sensitivity_manager.getlabel()).dump_info()
        print(filename)
        print(json.dumps(label_info, indent=4, ensure_ascii=False))
        document_manager.close_document()

    if document_manager:
        document_manager.quit()


def demo_set_sensitivity():
    """
    This demo function iterates through a list of filenames and their corresponding sensitivity labels.
    It copies a dummy document to each filename and sets the specified sensitivity label to the copied document.
    """
    for filename, sensitivity in [
        ("output/dummy_internal_use_only.xlsx", "InternalUseOnly"),
        ("output/dummy_commercial_in_confidence.xlsx", "CommercialInConfidence"),
        ("output/dummy_public.xlsx", "Public"),
        ("output/dummy_restricted_confidential.xlsx", "RestrictedConfidential"),
        ("output/dummy_restricted_sensitive.xlsx", "RestrictedSensitive"),
    ]:
        print(f"{filename = },{sensitivity = }")
        shutil.copy(
            "output/dummy.xlsx", filename
        )  # Copy a dummy file to the target filename.
        fullpath = os.path.join(os.getcwd(), filename)
        set_sensitivity_label_to_file(
            fullpath, sensitivity
        )  # Apply the sensitivity label.


def demo_initialize():
    """
    This function initializes the sensitivity labels definition file.
    It scans a directory for template documents, each named after a sensitivity label,
    and creates a JSON file mapping those labels to their properties.

    INITIALIZE sensitivity_configuration_file
    =========================================
    create with excel a few documents that serves as template for capturing sensitivity_label ids
    nb: it could be docx documents too

    As for example
    --------------
    C:.

    └───sensitivity_model

       CommercialInConfidence.xlsx

       InternalUseOnly.xlsx

       Public.xlsx

       RestrictedConfidential.xlsx

       RestrictedSensitive.xlsx

    """
    create_sensitivity_label_definition(
        extract_from=DEFAULT_SENSITIVITY_TEMPLATES,
        sensitivity_configuration_file=DEFAULT_SENSITIVITY_LABELS_DEFINITION,
    )


if __name__ == "__main__":
    # The following lines run the demo functions defined above.
    # Each function demonstrates a different aspect of managing sensitivity labels in Office documents.

    # First, initialize the sensitivity labels definition from predefined templates.
    demo_initialize()

    # Then, set sensitivity labels to some dummy documents based on the initialized definitions.
    demo_set_sensitivity()

    # Finally, show the label info for the documents to verify that the labels were correctly set.
    demo_show_label_info()

# Tips for Beginners:
# - Make sure to follow the setup instructions carefully.
# - Run this script in an environment where the office_toolbox package and its dependencies are correctly installed.
# - Modify the paths and filenames according to your setup.
# - Experiment by changing the sensitivity labels and observing how the script behaves.
# - Use print statements to debug and understand the flow of the script.
