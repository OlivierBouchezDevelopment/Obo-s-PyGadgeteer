# Sensitivity Label Management

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
