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
    set_sensitivity_label_to_document,
    DEFAULT_SENSITIVITY_TEMPLATES,
    DEFAULT_SENSITIVITY_LABELS_DEFINITION,
)


def demo_show_label_info():
    """
    This demo function iterates through a list of filenames and their corresponding sensitivity labels.
    It opens each document, retrieves its current sensitivity label, and prints the label information.

    note: the code between demo_excel_sensitivity_manager and demo_word_sensitivity_manager are redundant, for demo purpose,
          but in fact they are similar similar (the file extension makes the difference).
          classes and factory will adapt to a Word or an Excel document
    """
    document_manager = None
    for filename, sensitivity in [
        ("output/dummy_internal_use_only.docx", "InternalUseOnly"),
        ("output/dummy_commercial_in_confidence.docx", "CommercialInConfidence"),
        ("output/dummy_public.docx", "Public"),
        ("output/dummy_restricted_confidential.docx", "RestrictedConfidential"),
        ("output/dummy_restricted_sensitive.docx", "RestrictedSensitive"),
    ]:
        print(f"{filename = },{sensitivity = }")
        fullpath = os.path.abspath(filename)
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

    note: the code between demo_excel_sensitivity_manager and demo_word_sensitivity_manager are redundant, for demo purpose,
        but in fact they are similar similar (the file extension makes the difference).
        classes and factory will adapt to a Word or an Excel document
    """
    for filename, sensitivity in [
        ("output/dummy_internal_use_only.docx", "InternalUseOnly"),
        ("output/dummy_commercial_in_confidence.docx", "CommercialInConfidence"),
        ("output/dummy_public.docx", "Public"),
        ("output/dummy_restricted_confidential.docx", "RestrictedConfidential"),
        ("output/dummy_restricted_sensitive.docx", "RestrictedSensitive"),
    ]:
        print(f"{filename = },{sensitivity = }")
        shutil.copy(
            "output/dummy.docx", filename
        )  # Copy a dummy file to the target filename.
        fullpath = os.path.abspath(filename)
        set_sensitivity_label_to_file(
            absolute_path_to_filename=fullpath, sensitivity_label=sensitivity
        )  # Apply the sensitivity label.


def demo_create_document():
    """
    This demo create a document and set the Sensitivity Level

    note: the code between demo_excel_sensitivity_manager and demo_word_sensitivity_manager are redundant, for demo purpose,
        but in fact they are similar similar (the file extension makes the difference).
        classes and factory will adapt to a Word or an Excel document
    """
    # In fact, the code is exactly the same for a docx ... Just change the extension below
    filename = "output/dummy_new.docx"
    fullpath = os.path.abspath(filename)
    document_manager = document_manager_factory(fullpath)
    document_manager.create_document()
    set_sensitivity_label_to_document(document_manager, "InternalUseOnly")
    document_manager.close_document()
    document_manager = document_manager_factory(fullpath)
    if document_manager.document:
        sensitivity_manager = SensitivityLabelManager(document_manager.document)
        label_info = LabelInfoManager(sensitivity_manager.getlabel()).dump_info()
        print(filename)
        print(json.dumps(label_info, indent=4, ensure_ascii=False))
    else:
        print(f"Strange the file is not availble {filename} {fullpath =}")


def demo_initialize():
    """
    This function initializes the sensitivity labels definition file.
    It scans a directory for template documents, each named after a sensitivity label,
    and creates a JSON file mapping those labels to their properties.

    INITIALIZE sensitivity_configuration_file
    =========================================
    create with Word a few documents that serves as template for capturing sensitivity_label ids
    nb: it could be xlsx documents too
    You can initialize with excel or work, the configuration is common
    You have to initialize the first time, than the JSON file can be used for further call

    As for example
    --------------
    C:.

    └───sensitivity_model

       CommercialInConfidence.docx

       InternalUseOnly.docx

       Public.docx

       RestrictedConfidential.docx

       RestrictedSensitive.docx

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

    # Show creation of an empty document with Sensitivity Label
    demo_create_document()

# Tips for Beginners:
# - Make sure to follow the setup instructions carefully.
# - Run this script in an environment where the office_toolbox package and its dependencies are correctly installed.
# - Modify the paths and filenames according to your setup.
# - Experiment by changing the sensitivity labels and observing how the script behaves.
# - Use print statements to debug and understand the flow of the script.
