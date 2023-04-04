#!/bin/zsh

echo "Installing Spyglass..."
# Get the latest version of Spyglass
git clone https://github.com/Derpseh/Spyglass.git
cd Spyglass

# Check out the latest tag (the latest release)
# NOTE: Hardcoding this to v3.0.1 for now because releases after that
# are not considered stable. Leaving in my tag search code for now though.
# git checkout $(git describe --tags $(git rev-list --tags --max-count=1)) # Uncomment this line to get the latest release
git checkout 3.0.1

$py3 = python3 --version
# Check if python 3 is 3.9 or higher
if [[ $py3 == *"3.9"* || $py3 == *"3.10"* || $py3 == *"3.11"* ]]; then
    echo "Python 3.9 or higher detected."
else
    echo "Python 3.9 or higher not detected. Please install Python 3.9 or higher."
    exit 1
fi

# Check to see if the user wants a virtual environment
read -q "REPLY?Do you want to create a virtual environment? (y/n) "
echo
if [[ $REPLY =~ ^[Yy]$ ]]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
    source venv/bin/activate
fi

# Install requirements
pip install -r requirements.txt
