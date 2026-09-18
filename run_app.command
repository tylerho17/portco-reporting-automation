#!/bin/bash
# Double-click this file on a Mac to open the Board Pack Generator in your web browser.
# A Terminal window opens and must stay open while you use the app; close it to stop the app.

cd "$(dirname "$0")" || exit 1   # the project folder, wherever it has been copied to

# First run only: create the Python environment and install the packages (needs internet).
if [ ! -x .venv/bin/python ]; then
    echo "First run: setting up (this takes a few minutes)..."
    if ! command -v python3 >/dev/null; then
        echo "Python 3 isn't installed. Install it from https://www.python.org/downloads/ and try again."
        read -r -p "Press Return to close."
        exit 1
    fi
    python3 -m venv .venv && .venv/bin/pip install -r requirements.txt || {
        echo "Setup failed - see the messages above."
        read -r -p "Press Return to close."
        exit 1
    }
fi

# Open the page once the app has had a few seconds to start.
(sleep 3 && open "http://localhost:8501") &

echo "Starting the Board Pack Generator at http://localhost:8501 - close this window to stop it."
.venv/bin/streamlit run app.py
