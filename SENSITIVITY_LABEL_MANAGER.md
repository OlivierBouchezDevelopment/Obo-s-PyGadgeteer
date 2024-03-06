# Microsoft Information Protection (MIP)
Microsoft Information Protection (MIP) is a suite of solutions provided by Microsoft to help organizations discover, classify, protect, and monitor sensitive information across various platforms and environments. 

## Enhancing Document Security with Automated Sensitivity Labels in Python

In today's digital age, the concept of sensitivity labels plays a pivotal role in ensuring data security and compliance. However, when generating automated reports with Python, these reports lack sensitivity labels by default. This absence means that upon first opening a document, one must manually assign a sensitivity label— a step that introduces delays in the process of sharing script outputs with colleagues within an organization.

In search of a solution to streamline this process, I embarked on a journey that was anything but straightforward. Initially, I experimented with document automation using `pywin32`. This approach, while viable, fell short of my expectations, primarily because it necessitated having Excel installed—a requirement not always met, especially on backend systems or Linux servers.

My continued exploration led me to extensively use `openpyxl` for Excel document generation. Although `openpyxl` does not natively support sensitivity labeling, I discovered a workaround. By leveraging `custom_doc_props`—a list of custom properties within a workbook—I was able to manipulate the Microsoft Information Protection (MIP) information of an Excel document.

The foundation of both methodologies is the creation of a directory (`sensitivity_model`) containing a series of Excel documents that have been manually prepared with the required sensitivity labels. These documents serve as a template from which sensitivity label information is extracted, facilitating the generation of a JSON configuration. This setup process is a one-time requirement; subsequently, the JSON configuration can be employed to apply the correct sensitivity levels to documents programmatically.

Armed with this configuration, it becomes possible to dynamically assign sensitivity labels to documents. With `openpyxl`, we can generate and embed these labels directly into the document properties. Alternatively, for applications requiring `pywin32`, it offers the capability to interact with Excel or Word, directing these applications to modify or set the sensitivity level accordingly.

This dual-faceted approach not only enhances document security but also significantly optimizes the workflow of sharing sensitive information within an organization, bridging the gap between automation and compliance.


## Content:
Only for Excel:
[Sensitivity Label Management using openpyxl](#sensitivity-label-management-using-openpyxl)

For any office document (word, excel):
[Sensitivity Label Management using pywin32](#sensitivity-label-management-using-pywin32)

# Sensitivity Label Management using openpyxl

`pygadgeteer\openpyxl_toolbox\sensitivity_manager.py` intend to add sensitivity management feature to openpyxl. 
`openpyxl` is a Python library used to read and write Excel 2010 xlsx/xlsm/xltx/xltm files. openpyxl is especially useful because it doesn't require Microsoft Excel to be installed, making it an ideal choice for server-side processing of Excel files, automations, and data analysis tasks.

The module contains 
- `MSIP_Manager` a class to get or set sensitivity label to a openpyxl.Workbook
- `MSIP_Configuration` a class to handle a configuration file with sensitivity label definition. 
- create_sensitivity_label_definition a high level function that create a configuration file based on models (created with excel)
- `get_label_from_file` returns the sensitivity info from a file
- `set_label_to_workbook` set the sensitivity label to a workbook
- `set_label_to_file` set the sensitivity label to a file
  

In this tutorial, we'll learn how to programmatically create sensitivity labels, apply them to Excel workbooks, and verify their presence using the `openpyxl` library in Python. This process is essential for managing data sensitivity in automated workflows, ensuring that sensitive information is appropriately labeled and handled.

## Prerequisites

Before we start, ensure you have `openpyxl` installed in your Python environment. If not, you can install it using pip:

```shell
pip install openpyxl
```

## Step 1: Define Sensitivity Labels

First, we define our sensitivity labels. In a larger application, this might involve reading definitions from a file or database. For simplicity, we'll assume this step is abstracted by the `create_sensitivity_label_definition()` function.

## Step 2: Apply Labels to Workbooks

Next, we iterate through our defined labels, creating a new Excel workbook for each label and applying it accordingly. This demonstrates how to programmatically manage document sensitivity levels.

### Example Code

Here's a simplified script outline demonstrating the core functionality:

```python
from openpyxl import Workbook
from your_label_module import create_sensitivity_label_definition, get_label_from_file, set_label_to_workbook

if __name__ == "__main__":
    # Step 1: Create or load sensitivity label definitions
    msip_configuration = create_sensitivity_label_definition()

    # Step 2: Apply each label to a new workbook and save it
    for label_name in msip_configuration.labels():
        test_filename = f"output/dummy_openpyxl_{label_name}.xlsx"
        wb = Workbook()
        label = msip_configuration.get_sensitivity_label(label_name)
        set_label_to_workbook(wb, label)
        wb.save(test_filename)

        # Step 3: Verify the label was applied correctly
        check_label = get_label_from_file(test_filename)
        assert check_label is not None, "Label was not found in the workbook."
        assert label.LabelId == check_label.LabelId, "Label ID does not match."
        assert label.LabelName == check_label.LabelName, "Label Name does not match."
```

## Step 3: Verify Label Application

After saving the workbook with the applied label, we use `get_label_from_file()` to verify that the label was correctly applied. This is crucial for ensuring the integrity of your document labeling process.

## Conclusion

By following this tutorial, you've learned how to programmatically apply and verify sensitivity labels in Excel workbooks using `openpyxl`. This process is invaluable for automating data management tasks in applications that handle sensitive information.

Remember, while this example provides a basic framework, real-world applications may require more robust error handling and logging to ensure data integrity and application reliability.



# Sensitivity Label Management using pywin32

The Sensitivity Label Management feature is a key component of the Office Toolbox, designed to enhance document security and classification in organizations. Sensitivity Labels, such as "Public", "Internal Use Only", and "Commercial in Confidence", play a crucial role in document management. While libraries like openpyxl enable the creation of Excel documents via Python, managing Sensitivity Labels is not directly supported.

Obo's PyGadgeteer fills this gap by utilizing pywin32 and .NET communication, providing a powerful solution for applying Sensitivity Labels to Office documents. This feature is specifically tailored for Windows environments where Excel or Word is available, automating the label application process.

**Note:** COM is utilized as a last resort due to its potential for unpredictable behavior, including the occasional appearance of dialog or confirmation boxes. To mitigate this, documents are managed in a visible state, allowing for real-time adjustments based on application feedback.

## Getting Started

### Prerequisites

- Windows OS with Excel or Word installed.
- Python with the pywin32 library.

### Installation

To begin using Obo's PyGadgeteer, first ensure Python and the pywin32 library are installed on your system:

```bash
pip install pywin32
```

### Demonstrations

Included are two demonstration scripts that highlight how to apply Sensitivity Labels to Excel and Word documents:

- `demo_excel_sensitivity_manager`
- `demo_word_sensitivity_manager`

### Usage

Start by creating a `sensitivity_labels_definition.json` file to map Sensitivity Labels to their corresponding LabelId and LabelName:

```python
from office_toolbox.set_sensitivity_label import (
    create_sensitivity_label_definition,
    set_sensitivity_label_to_file
)
```

- **`create_sensitivity_label_definition`**: Prepares a configuration file linking "SensitivityLabel" with a LabelName and LabelId. This setup involves:
  1. Creating a folder (default: sensitivity_model).
  2. Using Word or Excel to generate documents with the required Sensitivity Labels.
  3. Executing the function without arguments for default settings.

    ```python
    create_sensitivity_label_definition()
    ```

- **`set_sensitivity_label_to_file`**: Assigns the chosen Sensitivity Label to a document. You must specify the document's full path.

    ```python
    sensitivity="InternalUseOnly"
    fullpath = os.path.abspath("filename.docx") # or "filename.xlsx"
    set_sensitivity_label_to_file(
        absolute_path_to_filename=fullpath, sensitivity_label=sensitivity
    )
    ```

### Architecture

At its core, `document_manager_factory` generates instances of `ExcelDocumentManager` or `WordDocumentManager`, each extending `AbstractDocumentManager`. These managers interface with the .NET framework to manipulate documents within the respective Excel or Word applications directly.
