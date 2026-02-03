#!/usr/bin/env python3
"""
Start PitchCraftAI web server
"""

import uvicorn
import webbrowser
import threading
import time

def open_browser():
    """Open browser after server starts"""
    time.sleep(1.5)
    webbrowser.open("http://localhost:8000")

if __name__ == "__main__":
    print("""
╔═══════════════════════════════════════════════════════════════╗
║                      PitchCraftAI MVP                         ║
║           AI-Powered DCF Model Generator                      ║
╚═══════════════════════════════════════════════════════════════╝
    """)
    print("Starting web server at http://localhost:8000")
    print("Press Ctrl+C to stop\n")

    # Open browser in background
    threading.Thread(target=open_browser, daemon=True).start()

    # Start server
    uvicorn.run(
        "web.api:app",
        host="0.0.0.0",
        port=8000,
        reload=False,
        log_level="info"
    )
