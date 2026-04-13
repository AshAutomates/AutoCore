Installation
============

AutoCore has been fully tested on Python 3.12. Using other versions may lead to compatibility issues with dependencies.

.. code-block:: bash

   pip install autocore

Linux Dependencies
------------------

After installing, run the following based on your distro:

.. code-block:: bash

   # Ubuntu/Debian
   sudo apt-get install wmctrl xdotool python3-tk xclip xdg-utils

   # RHEL/CentOS/Fedora
   sudo yum install wmctrl xdotool python3-tkinter xclip xdg-utils

.. list-table::
   :header-rows: 1
   :widths: 20 25 55

   * - Package
     - Used by
     - Purpose
   * - ``wmctrl``
     - ``window()``
     - List, focus, close, minimize, maximize, resize and move windows
   * - ``xdotool``
     - ``window()``, ``inspect()``
     - Minimize windows, restore them before resize/move, and get active window title
   * - ``python3-tk``
     - ``inspect()``
     - Render the Pixel Inspector GUI window
   * - ``xclip``
     - ``copy()``, ``inspect()``
     - Read and write clipboard content via pyperclip
   * - ``xdg-utils``
     - ``run()``
     - Open files with their default application via xdg-open

Chrome Installation
-------------------

AutoCore uses Chrome for browser automation. Install it before using ``browser()``.

**Windows:**

.. code-block:: bash

   winget install Google.Chrome

**Linux (Ubuntu/Debian/Mint):**

.. code-block:: bash

   wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
   sudo dpkg -i google-chrome-stable_current_amd64.deb
   sudo apt-get install -f -y

**Linux (RHEL/CentOS/Fedora):**

.. code-block:: bash

   wget https://dl.google.com/linux/direct/google-chrome-stable_current_x86_64.rpm
   sudo rpm -i google-chrome-stable_current_x86_64.rpm