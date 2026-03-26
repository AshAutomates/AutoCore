from setuptools import setup, find_packages
from setuptools.command.install import install
import platform

class PostInstall(install):
    def run(self):
        install.run(self)
        if platform.system() == "Linux":
            print("\n" + "="*60)
            print("AutoCore: Linux dependencies required for full functionality:")
            print("  Ubuntu/Debian:      sudo apt-get install wmctrl xdotool")
            print("  RHEL/CentOS/Fedora: sudo yum install wmctrl xdotool")
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
        "pyttsx3",
        "pyyaml",
    ],
    cmdclass={"install": PostInstall},
)