Quick Start
===========

This guide will get you up and running with AutoCore in minutes.

Basic Usage
-----------

All functions are available with a single import:

.. code-block:: python

   from autocore import *

Browser Automation
------------------

.. code-block:: python

   from autocore import *

   # Open browser and navigate to a URL
   dr = browser('https://example.com')

   # Click a button
   click(dr, 'id', 'login-button')

   # Type in a field
   write(dr, 'id', 'username', 'myuser')

   # Press enter
   press(dr, 'enter')

   # Close browser
   dr.quit()

Reading Screen Text
-------------------

.. code-block:: python

   from autocore import *

   # Read entire screen using OCR
   text = read()

   # Check if something is visible
   if 'error' in text:
       say("Error detected on screen")

Downloading Files
-----------------

.. code-block:: python

   from autocore import *

   dr = browser('https://example.com')

   # Click download button
   click(dr, 'id', 'download-button')

   # Wait for download to complete (300 seconds)
   filename = wait_download(300)
   print(f"Downloaded: {filename}")

   dr.quit()

Logging
-------

.. code-block:: python

   from autocore import *

   # Setup logging — terminal turns green on success, red on crash
   log_setup("my_script")

   dr = browser('https://example.com')
   click(dr, 'text', 'Submit')
   dr.quit()

   # All output is automatically saved to logs/ folder