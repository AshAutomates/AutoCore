from setuptools import setup, find_packages
from setuptools.command.install import install
import platform

class PostInstall(install):
    def run(self):
        install.run(self)

        print("\n┌" + "─" * 88 + "┐")
        print("│  AutoCore v1.0 — Installation Complete                                                │")
        print("│  Docs   : https://autocore.readthedocs.io                                             │")
        print("│  GitHub : https://github.com/AshAutomates/AutoCore                                    │")
        print("└" + "─" * 88 + "┘")

        if platform.system() == "Linux":
            print("\n  NOTE: Additional setup required on Linux\n")

            print("┌" + "─" * 88 + "┐")
            print("│  LINUX DEPENDENCIES                                                                    │")
            print("├" + "─" * 88 + "┤")
            print("│  To install, run:                                                                      │")
            print("│                                                                                        │")
            print("│  Ubuntu/Debian/Mint:                                                                   │")
            print("│    sudo apt-get install wmctrl xdotool python3-tk xclip xdg-utils                      │")
            print("│                                                                                        │")
            print("│  RHEL/CentOS/Fedora:                                                                   │")
            print("│    sudo yum install wmctrl xdotool python3-tkinter xclip xdg-utils                     │")
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

setup(
    name="autocore",
    version="1.0",
    author="Ash",
    url="https://github.com/AshAutomates/AutoCore",
    description="Automate Core Actions",
    long_description=open("README.rst").read(),
    long_description_content_type="text/x-rst",
    packages=find_packages(),
    install_requires=[
        "opencv-python",
        "numpy<2",
        "Pillow",
        "pyautogui",
        "pyperclip",
        "keyboard",
        "easyocr",
        "beautifulsoup4",
        "selenium",
        "webdriver-manager",
        "pywin32; sys_platform == 'win32'",
        "PyPDF2",
        "python-docx",
        "openpyxl",
        "python-pptx",
        "ebooklib",
        "odfpy",
        "extract-msg",
        "striprtf",
        "undetected-chromedriver",
        "requests",
        "pyttsx3",
        "pyyaml",
    ],
    cmdclass={"install": PostInstall},
)