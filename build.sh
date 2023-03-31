#!/bin/bash

echo "Creating Spyglass environment..."

read -p "Do you need to install Ubuntu python3-pip and python3-venv packages? (y/n) " -n 1 -r
echo
if [[ $REPLY =~ ^[Yy]$ ]]
then
    echo "Installing python3-venv, python3-pip, and upx..."
    echo "You may be asked to enter your password to continue."
    sudo apt update && sudo apt install -y python3-venv python3-pip upx
fi

echo "Creating a virtual environment..."
python3 -m venv venv
source venv/bin/activate

echo "Installing requirements..."
pip install -r requirements.txt

echo "Installing build tools..."
pip install pyinstaller

echo "Building..."
pyinstaller --clean spyglass.py -F -n Spyglass-Executable -c
chmod +x dist/Spyglass-Executable
mv dist/Spyglass .
echo "Done!"
echo "Cleaning up..."
rm -rf build/
rm -rf dist/

echo "Done! Run ./Spyglass-Executable to use Spyglass."