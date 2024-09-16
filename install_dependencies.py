import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Список пакетов 
packages = [
    "selenium",
    "webdriver-manager",
    "openpyxl"
]

for package in packages:
    try:
        __import__(package)
    except ImportError:
        print(f"Installing {package}...")
        install(package)
