from setuptools import setup, find_packages
from setuptools.command.install import install
import platform

class PostInstall(install):
    def run(self):
        install.run(self)

        print("\n" + "=" * 60)
        print("  AutoCore v1.0 installed successfully!")
        print("  Usage: from autocore import *")
        print("  GitHub:  https://github.com/AshAutomates/AutoCore")
        print("=" * 60)

        if platform.system() == "Linux":
            print("\n" + "!" * 60)
            print("!!                                                        !!")
            print("!!          LINUX DEPENDENCIES REQUIRED                   !!")
            print("!!                                                        !!")
            print("!!  Ubuntu/Debian/Mint:                                   !!")
            print("!!  sudo apt-get install wmctrl xdotool python3-tk        !!")
            print("!!                       xclip xdg-utils                  !!")
            print("!!                                                        !!")
            print("!!  RHEL/CentOS/Fedora:                                   !!")
            print("!!  sudo yum install wmctrl xdotool python3-tkinter       !!")
            print("!!                    xclip xdg-utils                     !!")
            print("!!                                                        !!")
            print("!" * 60)
            print("\n" + "!" * 60)
            print("!!                                                        !!")
            print("!!         CHROME REQUIRED FOR browser()                  !!")
            print("!!                                                        !!")
            print("!!  Ubuntu/Debian/Mint:                                   !!")
            print("!!  wget https://dl.google.com/linux/direct/              !!")
            print("!!       google-chrome-stable_current_amd64.deb           !!")
            print("!!  sudo dpkg -i google-chrome-stable_current_amd64.deb   !!")
            print("!!  sudo apt-get install -f -y                            !!")
            print("!!                                                        !!")
            print("!!  RHEL/CentOS/Fedora:                                   !!")
            print("!!  wget https://dl.google.com/linux/direct/              !!")
            print("!!       google-chrome-stable_current_x86_64.rpm          !!")
            print("!!  sudo rpm -i google-chrome-stable_current_x86_64.rpm   !!")
            print("!!                                                        !!")
            print("!" * 60)

        if platform.system() == "Windows":
            print("\n" + "!" * 60)
            print("!!                                                        !!")
            print("!!         CHROME REQUIRED FOR browser()                  !!")
            print("!!                                                        !!")
            print("!!  Run: winget install Google.Chrome                     !!")
            print("!!                                                        !!")
            print("!" * 60)

        print("\n" + "=" * 60 + "\n")

setup(
    name="autocore",
    version="1.0",
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