# Sensitivity Label Management

## Current Features

- **Sensitivity Label Management**: A feature designed to help manage and apply sensitivity labels to your data, ensuring appropriate handling and protection according to its classification. [SensitivityLabelManager.md](doc:SensitivityLabelManager.md)

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

## License

Specify your project's license here.

---

This README.md file provides a concise overview of your project, how to get started, and how to contribute. Adjust the sections as your project evolves or as you add more features.