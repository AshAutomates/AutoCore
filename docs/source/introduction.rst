Introduction
============

Automate Core Actions

.. image:: https://img.shields.io/badge/version-1.0-blue
   :alt: Version

.. image:: https://img.shields.io/badge/python-3.12-blue
   :alt: Python

.. image:: https://img.shields.io/badge/platform-Windows%20%7C%20Linux-blue
   :alt: Platform

A Python library that combines GUI automation, headless browser control and common actions into a single library.

With a single import, we get everything we need to automate repetitive tasks on Windows and Linux, no juggling multiple libraries or complex setup.

----

Security & Privacy
------------------

AutoCore runs entirely on your local machine. No data is sent to any external server at any point.

This makes it suitable for:

- On-premise deployments with strict data security policies.
- Automation involving sensitive or confidential data.
- Compliance-sensitive industries like finance, healthcare and legal.



----

Functions
---------

.. list-table::
   :header-rows: 1
   :widths: 30 70

   * - Function
     - Description
   * - ``browser(url)``
     - Open Chrome browser and navigate to URL
   * - ``click(...)``
     - Click on image, text, coordinates, color or web element
   * - ``click_right(...)``
     - Right-click on image, text, coordinates, color or web element
   * - ``copy(...)``
     - Copy text from active window, clipboard, coordinates or web element
   * - ``csv_to_xlsx(...)``
     - Convert CSV file to XLSX
   * - ``date()``
     - Current day of month (1-31)
   * - ``day()``
     - Current day of week (monday, tuesday, ...)
   * - ``drag(...)``
     - Drag from source to target
   * - ``dropdown_select(...)``
     - Select item from a dropdown
   * - ``erase(...)``
     - Clear text from input fields
   * - ``find_browser(...)``
     - Find text in browser using Ctrl+F
   * - ``find_key(data, key)``
     - Recursively find all values of a key in nested data
   * - ``find_str(...)``
     - Extract substring between two markers
   * - ``hour()``
     - Current hour (0-23)
   * - ``inspect()``
     - GUI tool to inspect pixel position and color (Windows only)
   * - ``log_setup(title)``
     - Setup logging with terminal color status
   * - ``minute()``
     - Current minute (0-59)
   * - ``month()``
     - Current month (1-12)
   * - ``press(...)``
     - Press keyboard keys
   * - ``read(...)``
     - Read text from screen or browser via OCR or extract text from files
   * - ``run(item)``
     - Run a file or application
   * - ``say(text)``
     - Speak text using offline Text-to-Speech
   * - ``screenshot(...)``
     - Take a screenshot of full screen or region
   * - ``scroll(...)``
     - Scroll up, down, to top or to bottom
   * - ``second()``
     - Current second (0-59)
   * - ``wait(...)``
     - Wait with countdown, wait for element or wait for color
   * - ``wait_download(...)``
     - Monitor downloads folder for completion or download directly via URL
   * - ``window(...)``
     - List, focus, close, minimize, maximize, resize or move windows
   * - ``write(...)``
     - Type text in active window or web element
   * - ``year()``
     - Current year
   * - ``zoom(...)``
     - Zoom in/out by steps or set zoom percentage

File Formats supported by ``read()``
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

PDF, DOCX, PPTX, ODT, RTF, CSV, TSV, XLSX, SQLite, JSON, YAML, XML, INI/CFG, TXT, LOG, MD, HTML, EML, MSG, EPUB, SH, BAT, PY

----

Platform Support
----------------

Supported: Windows, Linux

Not Supported: macOS

----

Quick Example
-------------

.. code-block:: python

   from autocore import *

   # Log setup with auto color change on success/failure
   log_setup("demo_script")

   # Open browser and click a button
   dr = browser('https://example.com')
   click(dr, 'id', 'login-button')

   # Write and press keys on the initiated browser
   write(dr, 'id', 'username', 'myuser')
   press(dr, 'enter')

   # OCR - read text from screen
   text = read()
   if 'error' in text:
       say("Error detected on screen")

   # Click on button with download text on it
   click('Download')

   # Wait for download to complete
   wait_download(10)