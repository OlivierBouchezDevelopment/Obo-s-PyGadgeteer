# Sensitivity Label Management
A notable feature within the Office Toolbox is the management of Sensitivity Labels for Office documents. Sensitivity Labels are widely used in various organizations to classify documents with tags such as "Public", "Internal Use Only", and "Commercial in Confidence". While tools like openpyxl facilitate the creation of Excel documents from Python, they do not inherently support the management of Sensitivity Labels.

Obo's PyGadgeteer bridges this gap by leveraging pywin32 and .NET communication, offering a robust solution for applying Sensitivity Labels to Office documents. This feature is designed to work on Windows platforms with Excel or Word installed, streamlining the document management process by automating the application of Sensitivity Labels.

Note:
    Using COM is not my prefered method, 
    I'm using it when there is no other solution. COM could have unforeseen result as it depends on the application. Sometimes a dialog or confirmation box can appear.
    I always use the application Visible state, so I can adjust when I see the dialog message from the application. 

## Getting Started

### Prerequisites
Windows OS with Excel or Word installed.
Python with pywin32 library.

### Installation
To use Obo's PyGadgeteer, ensure you have Python installed on your system along with the pywin32 library:

```bash
pip install pywin32
```

### Demonstrations
Two demonstration scripts are provided to showcase the functionality of our Sensitivity Label management:

- `demo_excel_sensitivity_manager`
- `demo_word_sensitivity_manager`
  
These demos illustrate how to apply Sensitivity Labels to Excel and Word documents, respectively.

### Usage
First, create a *sensitivity_labels_definition.json* file to map each Sensitivity Label to a corresponding LabelId and LabelName:

```python
from office_toolbox.set_sensitivity_label import (
    create_sensitivity_label_definition,
    set_sensitivity_label_to_file
)
```

- **`create_sensitivity_label_definition`**: 
To use the sensitivity manager, you need to create a configuration file. This file associates to each "SensitivityLabel" a LabelName and a LabelId. 
`create_sensitivity_label_definition` will create the configuration files. 
1. Create a folder (by default sensitivity_model)
2. Use Word or Excel to create a documents with the SensitivityLabel you need
3. Call the function without argument if you use the default
   ```python
   create_sensitivity_label_definition()
   ``` 
   

- **`set_sensitivity_label_to_file`**: Applies the specified Sensitivity Label to a document.
  You have to provide the full path to the file
    ```python
    sensitivity="InternalUseOnly"
    fullpath = os.path.abspath("filename.docx") # "filename.xlsx"
    set_sensitivity_label_to_file(
        absolute_path_to_filename=fullpath, sensitivity_label=sensitivity
    )  # Apply the sensitivity label.

    ```

### Architecture
Under the hood, document_manager_factory creates instances of ExcelDocumentManager or WordDocumentManager, each an implementation of AbstractDocumentManager. These managers are designed to interact with the .NET framework, directly manipulating documents within Excel or Word applications.


