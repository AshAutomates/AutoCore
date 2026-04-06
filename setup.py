from setuptools import setup, find_packages
from setuptools.command.install import install
import platform

class PostInstall(install):
    def run(self):
        install.run(self)
        print("\n" + "="*60)
        print("AutoCore v0.1.0 installed successfully!")
        print("Usage: from autocore import *")
        print("Docs:  https://github.com/AshAutomates/AutoCore")
        if platform.system() == "Linux":
            print("\n" + "!"*60)
            print("!! IMPORTANT: Linux dependencies required !!")
            print("!! Run the following commands to ensure full functionality !!")
            print("!!")
            print("!!  Ubuntu/Debian:      sudo apt-get install wmctrl xdotool python3-tk")
            print("!!  RHEL/CentOS/Fedora: sudo yum install wmctrl xdotool python3-tkinter")
            print("!"*60)
        print("="*60 + "\n")

setup(
    name="autocore",
    version="0.1.0",
    author="Ash",
    url="https://github.com/AshAutomates/AutoCore",
    description="Automate Core Actions",
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
        "requests",
        "pyttsx3",
        "pyyaml",
    ],
    cmdclass={"install": PostInstall},
)