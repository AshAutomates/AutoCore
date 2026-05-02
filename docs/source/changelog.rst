Changelog
=========

Version 1.4
-----------

*Released: May 2026*

- ``click()`` : Added JS fallback when element click is intercepted by overlays or ads.
- ``click()`` : Scroll element into view before clicking to handle off-viewport elements.
- ``click_right()`` : Scroll element into view before right-clicking to handle off-viewport elements.
- Fixed missing ``piper-tts`` and ``huggingface_hub`` from ``install_requires`` in ``setup.py``.
- Changed comment style from ``#=====`` to ``#-----abc-----`` across the library.


Version 1.3
-----------

*Released: May 2026*

- ``say()`` : Switched from pyttsx3 to Piper TTS (offline neural TTS) using ``en_US-libritts_r-medium`` voice model.
- ``browser()`` : Explicitly set Chrome download directory and exposed it as ``driver.download_dir``.
- Improved audio engine check at import where it verifies ``aplay`` availability and real audio hardware.
- Removed unused dependencies ``pyttsx3`` & ``keyboard``.
- Added ``espeak-ng`` and ``alsa-utils`` to Linux dependencies.
- Added ``__version__`` in ``__init__.py`` to support ``autocore.__version__`` checks.
- Added ``quick_example.py`` in tests folder.


Version 1.2
-----------

*Released: April 2026*

- Fixed ``pyautogui.FAILSAFE`` crash on Linux import.
- Fixed ``NameError`` in ``browser()`` when Chrome version detection fails.
- Added PyPI downloads badge.
- Added ``RELEASE.md`` with version bump checklist and release steps.


Version 1.1
-----------

*Released: April 2026*

- ``browser()`` : Fixed ChromeDriver version mismatch when Chrome update is delayed.
- Fixed post-install message not shown when installed via ``uv`` or other modern package managers.
- Fixed ``import autocore`` crash on Linux environments with no display server.
- Added ``FORCE_JAVASCRIPT_ACTIONS_TO_NODE24: true`` to workflow.yml to suppress Node.js 20 deprecation warnings.
- Added build badge to README.
- Added ReadTheDocs badge to README.
- Centered logo in README.
- Done minor changes to AutoCore logo.


Version 1.0
-----------

*Released: April 2026*

Initial release of AutoCore.

**Functions included:**

- ``browser()`` : Open Chrome browser and navigate to URL
- ``click()`` : Click on image, text, coordinates, color, or web element
- ``click_right()`` : Right-click on image, text, coordinates, color, or web element
- ``copy()`` : Copy text from active window, clipboard, coordinates, or web element
- ``csv_to_xlsx()`` : Convert CSV file to XLSX
- ``date()`` : Current day of month
- ``day()`` : Current day of week
- ``drag()`` : Drag from source to target
- ``dropdown_select()`` : Select item from a dropdown
- ``erase()`` : Clear text from input fields
- ``find_browser()`` : Find text in browser using Ctrl+F
- ``find_key()`` : Recursively find all values of a key in nested data
- ``find_str()`` : Extract substring between two markers
- ``hour()`` : Current hour
- ``inspect()`` : GUI tool to inspect pixel position and color (Windows only)
- ``log_setup()`` : Setup logging with terminal color status
- ``minute()`` : Current minute
- ``month()`` : Current month
- ``press()`` : Press keyboard keys
- ``read()`` : Read text from screen or browser via OCR or extract text from files
- ``run()`` : Run a file or application
- ``say()`` : Speak text using offline Text-to-Speech
- ``screenshot()`` : Take a screenshot of full screen or region
- ``scroll()`` : Scroll up, down, to top, or to bottom
- ``second()`` : Current second
- ``wait()`` : Wait with countdown, wait for element, or wait for color
- ``wait_download()`` : Monitor downloads folder for completion or download directly via URL
- ``window()`` : List, focus, close, minimize, maximize, resize, or move windows
- ``write()`` : Type text in active window or web element
- ``year()`` : Current year
- ``zoom()`` : Zoom in/out by steps or set zoom percentage