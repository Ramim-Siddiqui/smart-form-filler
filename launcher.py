"""Starts the next free copy of Smart Form Filler, so several websites can be filled at the same time.

Copy N uses screen port 8500+N and browser port 9221+N (see engine.SLOT).
"""
import os
import socket
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
MAX_COPIES = 9


def port_free(port):
    with socket.socket() as s:
        return s.connect_ex(("127.0.0.1", port)) != 0


# The first copy whose screen is not running. Its browser may still be open from before; it reconnects to it.
slot = next((n for n in range(1, MAX_COPIES + 1) if port_free(8500 + n)), None)
if slot is None:
    print(f"All {MAX_COPIES} copies are already running. Close one first.")
    input("Press Enter to close...")
    sys.exit(1)

print(f"Starting Smart Form Filler - copy {slot} (keep this window open while you use it)")
if slot > 1:
    print("This copy has its own browser window and login, separate from the other copies.")
os.environ["SFF_SLOT"] = str(slot)
sys.exit(subprocess.call([sys.executable, "-m", "streamlit", "run", os.path.join(HERE, "app.py"),
                          "--server.port", str(8500 + slot), "--browser.gatherUsageStats", "false"], cwd=HERE))
