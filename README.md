# Obo's PyGadgeteer

Welcome to the Obo's PyGadgeteer repository, a versatile collection of features designed to address a broad spectrum of specific concerns. Unlike traditional libraries, this project does not aim for distribution as a pip package. Instead, it serves as a rich repository where developers can find and reuse code snippets across various scenarios. Our goal is to provide a diverse toolkit where each feature stands independently, catering to different needs without the necessity for interconnected functionality.

## Current Features

- **Sensitivity Label Management**: A feature designed to help manage and apply sensitivity labels to your data, ensuring appropriate handling and protection according to its classification. [SensitivityLabelManager](SENSITIVITY_LABEL_MANAGER.md)

## Getting Started

To get started with our project, we recommend the following steps:

1. **Clone the Repository**: Clone this repository to your local machine to have access to all the available features.

    ```
    git clone <repository-url>
    ```

2. **Explore the Code**: Dive into the individual features within the repository. Each feature is designed to be self-contained, so you can easily understand and extract the code you need for your own projects.

3. **Utilize the Documentation**: Detailed documentation is embedded within the codebase, employing docstrings to explain the functionality and usage of each feature. To generate and view this documentation locally, follow the steps below:

    - **Build HTML Documentation**: Use the `build_docs.bat` script to generate HTML documentation from the codebase's docstrings.

        ```
        .\script\build_docs.bat
        ```

    - **Serve Documentation Locally**: To view the generated documentation in your browser, start a local documentation server with `doc_server.bat`.

        ```
        .\script\doc_server.bat
        ```

4. **Optional: Build Package**: While the project is not intended for pip distribution, you can still package it into a wheel file using `build_package.bat` if needed for specific scenarios.

    ```
    .\script\build_package.bat
    ```

## Contributing

Your contributions are welcome! If you have ideas for new features or improvements, please feel free to fork the repository, make your changes, and submit a pull request.

## MIT License

Copyright (c) [2024] [Olivier Bouchez]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

