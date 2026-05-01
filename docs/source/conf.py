import os
import sys
sys.path.insert(0, os.path.abspath('../..'))

project = 'AutoCore'
copyright = '2026, Ash'
author = 'Ash'
release = '1.3'

extensions = [
    'sphinx.ext.autodoc',
    'sphinx.ext.napoleon',
    'sphinx.ext.viewcode',
]

templates_path = ['_templates']
exclude_patterns = []
add_module_names = False

html_theme = 'furo'
html_static_path = ['_static']
html_logo = "_static/logo.png"
html_css_files = ['custom.css']

html_sidebars = {
    "**": [
        "sidebar/scroll-start.html",
        "sidebar/brand.html",
        "sidebar/search.html",
        "sidebar/navigation.html",
        "sidebar/scroll-end.html",
    ]
}

html_theme_options = {
    "navigation_with_keys": True,
    "sidebar_hide_name": True,
}
# Custom sections to avoid rendering of those as plain unstyled text.
napoleon_custom_sections = [
    'Modes',
    'Output',
    'Platform',
    'Usage',
    'Features',
    'Log file format',
    'Session numbering',
    'Value Logic',
    'Supported file formats',
]