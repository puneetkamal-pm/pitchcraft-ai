#!/usr/bin/env python3
"""
Start PitchCraftAI web server
"""

import uvicorn
import webbrowser
import threading
import time
import sys
from pathlib import Path

# Ensure repo root is on path so pitchcraft package resolves
sys.path.insert(0, str(Path(__file__).parent))

def open_browser():
    """Open browser after server starts"""
    time.sleep(1.5)
    webbrowser.open("http://localhost:8001")

if __name__ == "__main__":
    print("""
╔═══════════════════════════════════════════════════════════════╗
║                      PitchCraftAI MVP                         ║
║           AI-Powered DCF Model Generator                      ║
╚═══════════════════════════════════════════════════════════════╝
    """)
    print("Starting web server at http://localhost:8001")
    print("Press Ctrl+C to stop\n")

    # Open browser in background
    threading.Thread(target=open_browser, daemon=True).start()

    # Start server
    uvicorn.run(
        "pitchcraft.web.api:app",
        host="0.0.0.0",
        port=8001,
        reload=False,
        log_level="info"
    )
