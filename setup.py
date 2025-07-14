"""
城北中央公園テニスコート予約システム用セットアップスクリプト
py2appを使用してMac用アプリケーションを作成します
"""
from setuptools import setup

APP = ['johoku_app.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'site_packages': True,
    'includes': ['PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'pandas', 'selenium', 'webdriver_manager'],
    'excludes': ['tkinter', 'matplotlib', 'scipy', 'PyInstaller', 'test', 'unittest'],
    'resources': [],
    'optimize': 0,
    'compressed': False,
    'plist': {
        'CFBundleName': '城北中央公園テニスコート予約',
        'CFBundleDisplayName': '城北中央公園テニスコート予約',
        'CFBundleGetInfoString': '城北中央公園テニスコート予約システム',
        'CFBundleIdentifier': 'com.yourcompany.johokuapp',
        'CFBundleVersion': '1.0.2',
        'CFBundleShortVersionString': '1.0.2',
        'NSHumanReadableCopyright': '© 2025',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.9',
    },
}

setup(
    name='城北中央公園テニスコート予約',
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)