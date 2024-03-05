# Obo's PyGadgeteer

## Overview

Obo's PyGadgeteer is a versatile Python library designed to aggregate a variety of modules and classes to support an extensive range of use cases. This project serves as a comprehensive toolbox, offering both reusable code snippets and fully integrated solutions for diverse projects. Our goal is to provide a collection of tools that deliver general functionalities, enhancing productivity and efficiency for developers working across different domains.

## Features

### Office Toolbox
Our Office Toolbox offers a suite of routines aimed at simplifying interactions with Office documents. This includes, but is not limited to, creating, manipulating, and managing sensitivity labels within Office files.

#### Sensitivity Label Management
A notable feature within the Office Toolbox is the management of Sensitivity Labels for Office documents. Sensitivity Labels are widely used in various organizations to classify documents with tags such as "Public", "Internal Use Only", and "Commercial in Confidence". While tools like openpyxl facilitate the creation of Excel documents from Python, they do not inherently support the management of Sensitivity Labels.

Obo's PyGadgeteer bridges this gap by leveraging pywin32 and .NET communication, offering a robust solution for applying Sensitivity Labels to Office documents. This feature is designed to work on Windows platforms with Excel or Word installed, streamlining the document management process by automating the application of Sensitivity Labels.

### Getting Started

#### Prerequisites
Windows OS with Excel or Word installed.
Python with pywin32 library.

#### Installation
To use Obo's PyGadgeteer, ensure you have Python installed on your system along with the pywin32 library:

```bash
pip install pywin32
```

#### Demonstrations
Two demonstration scripts are provided to showcase the functionality of our Sensitivity Label management:

- `demo_excel_sensitivity_manager`
- `demo_word_sensitivity_manager`
  
These demos illustrate how to apply Sensitivity Labels to Excel and Word documents, respectively.

#### Usage
First, create a *sensitivity_labels_definition.json* file to map each Sensitivity Label to a corresponding LabelId and LabelName:

```python
from office_toolbox.set_sensitivity_label import (
    create_sensitivity_label_definition,
    set_sensitivity_label_to_file
)
```

- `create_sensitivity_label_definition`: Scans a directory of "template" files to generate a configuration file mapping labels to their properties.
- `set_sensitivity_label_to_file`: Applies the specified Sensitivity Label to a document.

#### Architecture
Under the hood, document_manager_factory creates instances of ExcelDocumentManager or WordDocumentManager, each an implementation of AbstractDocumentManager. These managers are designed to interact with the .NET framework, directly manipulating documents within Excel or Word applications.

## Future Directions
We are committed to continuously enhancing Obo's PyGadgeteer with new features and improvements. Feedback and contributions from the community are highly welcomed as we strive to build a robust and versatile toolbox for developers.

