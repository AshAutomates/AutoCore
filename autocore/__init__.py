__version__ = "1.4"
import os
import platform

def _show_install_info():

    print("\n┌" + "─" * 88 + "┐")
    print("│  AutoCore v1.4 — Installation Complete                                                 │")
    print("│  Docs   : https://autocore.readthedocs.io                                              │")
    print("│  GitHub : https://github.com/AshAutomates/AutoCore                                     │")
    print("└" + "─" * 88 + "┘")

    if platform.system() == "Linux":
        print("\n  NOTE: Additional setup required on Linux\n")

        print("┌" + "─" * 88 + "┐")
        print("│  LINUX DEPENDENCIES                                                                    │")
        print("├" + "─" * 88 + "┤")
        print("│  To install, run:                                                                      │")
        print("│                                                                                        │")
        print("│  Ubuntu/Debian/Mint:                                                                   │")
        print("│    sudo apt-get install wmctrl xdotool python3-tk xclip xdg-utils espeak-ng            │")
        print("│                                                                                        │")
        print("│  RHEL/CentOS/Fedora:                                                                   │")
        print("│    sudo yum install wmctrl xdotool python3-tkinter xclip xdg-utils espeak-ng           │")
        print("└" + "─" * 88 + "┘")

        print("\n┌" + "─" * 88 + "┐")
        print("│  CHROME — Required for browser() to work                                               │")
        print("├" + "─" * 88 + "┤")
        print("│  To install, run:                                                                      │")
        print("│                                                                                        │")
        print("│  Ubuntu/Debian/Mint:                                                                   │")
        print("│    wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb      │")
        print("│    sudo dpkg -i google-chrome-stable_current_amd64.deb                                 │")
        print("│    sudo apt-get install -f -y                                                          │")
        print("│                                                                                        │")
        print("│  RHEL/CentOS/Fedora:                                                                   │")
        print("│    wget https://dl.google.com/linux/direct/google-chrome-stable_current_x86_64.rpm     │")
        print("│    sudo rpm -i google-chrome-stable_current_x86_64.rpm                                 │")
        print("└" + "─" * 88 + "┘")

    if platform.system() == "Windows":
        print("\n  NOTE: Additional setup required on Windows\n")

        print("┌" + "─" * 88 + "┐")
        print("│  CHROME — Required for browser() to work                                               │")
        print("├" + "─" * 88 + "┤")
        print("│  To install, run:                                                                      │")
        print("│                                                                                        │")
        print("│    winget install Google.Chrome                                                        │")
        print("└" + "─" * 88 + "┘")

    print()


_flag = os.path.join(os.path.dirname(__file__), ".install_shown")
if not os.path.exists(_flag):
    _show_install_info()
    open(_flag, "w").close()

from ._lib import *
