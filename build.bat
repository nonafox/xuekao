@echo off
python build.py
pyinstaller -F -w index.py
