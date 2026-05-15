#!/usr/bin/env bash
# GHG Excel QA Tool — Mac / Linux launcher
# Run: bash run_app.sh   (or chmod +x run_app.sh && ./run_app.sh)

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" && pwd)"

# Prefer the project venv if it exists
if [ -f "$SCRIPT_DIR/.venv/bin/streamlit" ]; then
    STREAMLIT="$SCRIPT_DIR/.venv/bin/streamlit"
else
    STREAMLIT="streamlit"
fi

echo "Starting GHG QA Tool at http://localhost:8501 …"
"$STREAMLIT" run "$SCRIPT_DIR/app/streamlit_app.py" \
    --server.headless true \
    --server.port 8510 \
    --browser.serverAddress localhost \
    --browser.gatherUsageStats false
