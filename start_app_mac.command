#!/bin/bash
cd "$(dirname "$0")"

echo "==================================================="
echo "    Master Inspection Plan — Starting Server"
echo "==================================================="
echo ""

if command -v python3 >/dev/null 2>&1; then
    python3 app.py inspection_plan.db
elif [ -f "/usr/bin/python3" ]; then
    /usr/bin/python3 app.py inspection_plan.db
else
    echo "ERROR: Python 3 not found on this Mac."
    echo "Please install Python 3 or run via Command Line Tools."
    read -p "Press Enter to exit..."
fi
