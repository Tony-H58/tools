# -*- coding: utf-8 -*-
"""
Color Comments Record_QA Launcher
Automatically start Flask server + open HTML page
"""

import subprocess
import webbrowser
import os
import sys
import time
import signal
from pathlib import Path

def main():
    # Get current directory
    script_dir = Path(__file__).parent
    excel_file = script_dir / 'Color Comments Record_QA.xlsx'
    html_file = script_dir / 'Color Comments Record_QA.html'
    app_file = script_dir / 'app.py'

    print("\n" + "="*60)
    print("Color Comments Record_QA - Startup System")
    print("="*60)

    # Check files
    if not excel_file.exists():
        print(f"ERROR: {excel_file.name} not found")
        input("Press Enter to exit...")
        return

    if not html_file.exists():
        print(f"ERROR: {html_file.name} not found")
        input("Press Enter to exit...")
        return

    if not app_file.exists():
        print(f"ERROR: {app_file.name} not found")
        input("Press Enter to exit...")
        return

    print("OK - Files verified")
    print()

    # Check Python packages
    print("Checking required packages...")
    try:
        import flask
        import openpyxl
        print("OK - Packages installed")
    except ImportError:
        print("Installing required packages for first use...")
        os.system(f"{sys.executable} -m pip install flask openpyxl -q")
        print("OK - Packages installed")

    print()

    # Start Flask server (background)
    print("Starting Flask server...")
    flask_process = subprocess.Popen(
        [sys.executable, str(app_file)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd=str(script_dir),
        universal_newlines=True
    )

    # Wait for server startup
    time.sleep(3)

    if flask_process.poll() is not None:
        # Flask exited, show error
        stdout, stderr = flask_process.communicate()
        print("ERROR - Flask server failed to start")
        print("\nError message:")
        print(stderr if stderr else stdout)
        input("Press Enter to exit...")
        return

    print("OK - Flask server started")
    print()

    # Open HTML page
    print("Opening HTML page...")
    try:
        # Open Flask server (NOT file://)
        html_url = "http://localhost:5000"

        # Wait for server to fully start
        time.sleep(2)

        # Try to open browser
        if webbrowser.open(html_url):
            print("OK - HTML page opened")
        else:
            print("WARNING - Could not open browser, please visit manually:")
            print(f"   {html_url}")
    except Exception as e:
        print(f"WARNING - Browser open failed: {e}")
        print(f"   Please visit: http://localhost:5000")
    print()

    # Show status
    print("="*60)
    print("OK - System started")
    print("="*60)
    print()
    print("Usage:")
    print("  1. Edit data in the opened webpage")
    print("  2. Click 'Save' when done")
    print("  3. Close window to exit")
    print()
    print("URL: http://localhost:5000")
    print()
    print("-" * 60)
    print("Press Ctrl+C or close window to stop server...")
    print("-" * 60)
    print()

    try:
        # Keep running until user interrupt
        flask_process.wait()
    except KeyboardInterrupt:
        print("\nStopping server...")
        flask_process.terminate()
        flask_process.wait(timeout=5)
        print("OK - Server stopped")
    except Exception as e:
        print(f"ERROR: {e}")
    finally:
        # Make sure server is closed
        if flask_process.poll() is None:
            flask_process.kill()

if __name__ == '__main__':
    main()
