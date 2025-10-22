#!/usr/bin/env bash
set -e  # exit immediately on error

echo "=== Setting up Python virtual environment for Report Generator ==="

# Step 1: Check Python
if ! command -v python3 &> /dev/null; then
    echo "Python3 not found! Please install it first."
    exit 1
fi

# Step 2: Create and activate venv
if [ ! -d "myenv" ]; then
    python3 -m venv myenv
    echo "✅ Virtual environment created."
else
    echo "ℹ️  Virtual environment already exists."
fi

source myenv/bin/activate
echo "✅ Virtual environment activated."

# Step 3: Upgrade pip and install requirements
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt

# Step 4: Verify external converters (Linux only)
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    if command -v libreoffice &> /dev/null; then
        echo "✅ LibreOffice found."
    else
        echo "⚠️  LibreOffice not found. Please install it if you want PDF conversion fallback:"
        echo "    sudo apt install libreoffice"
    fi

    if command -v onlyoffice-desktopeditors &> /dev/null; then
        echo "✅ ONLYOFFICE found."
    else
        echo "⚠️  ONLYOFFICE not found. Install via .deb package for best PDF quality:"
        echo "    wget https://download.onlyoffice.com/install/desktop/editors/linux/onlyoffice-desktopeditors_amd64.deb"
        echo "    sudo apt install ./onlyoffice-desktopeditors_amd64.deb"
    fi
fi

echo "=== Setup complete! ==="
echo "To activate your environment manually, run:"
echo "source myenv/bin/activate"