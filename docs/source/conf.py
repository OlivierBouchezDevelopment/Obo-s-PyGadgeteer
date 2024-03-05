# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = "Obo's PyGadgeteer"
copyright = "2024, Olivier Bouchez"
author = "Olivier Bouchez"
release = "0.0.1"

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = []

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]


# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

# Set the html_theme to 'sphinx_rtd_theme'
html_theme = "sphinx_rtd_theme"

html_static_path = ["_static"]
extensions = [
    "sphinx.ext.autodoc",
    "myst_parser",
    "sphinx_rtd_theme",
    # other extensions...
]
