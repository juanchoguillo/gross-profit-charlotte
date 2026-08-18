#!/bin/bash
# Double-click this file (macOS Finder) to launch the Gross Profit dashboard.
cd "$(dirname "$0")" || exit 1

# First-time setup: create the virtual environment and install dependencies.
if [ ! -d ".venv" ]; then
  echo "First-time setup (about a minute)…"
  python3 -m venv .venv || { echo "Could not create environment. Is Python 3 installed?"; read -r; exit 1; }
  ./.venv/bin/pip install --quiet --upgrade pip
  ./.venv/bin/pip install --quiet -r requirements.txt
fi

# Skip Streamlit's first-run "enter your email" prompt so the app opens directly.
mkdir -p "$HOME/.streamlit"
if [ ! -f "$HOME/.streamlit/credentials.toml" ]; then
  printf '[general]\nemail = ""\n' > "$HOME/.streamlit/credentials.toml"
fi

echo "Starting the dashboard — your browser will open at http://localhost:8501"
echo "(Leave this window open while you use it. Close it or press Ctrl-C to stop.)"
exec ./.venv/bin/streamlit run app.py --server.gatherUsageStats false
