#!/bin/bash

echo "Creating virtual environment"
python3 -m venv venv
source venv/bin/activate
echo "Installing requirements"
pip install -r requirements.txt
echo "Installing build tools"
pip install pyinstaller
echo "Building"
pyinstaller --clean Spyglass-cli.py -F -n Spyglass -c
mv dist/Spyglass .
echo "Done"
echo "Cleaning up"
rm -rf build/
rm -rf dist/