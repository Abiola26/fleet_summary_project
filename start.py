import sys
import subprocess

# Use subprocess to run streamlit via the current python interpreter
# This avoids issues where 'streamlit' is not in the system PATH
try:
    cmd = [sys.executable, "-m", "streamlit", "run", "fleet_dashboard.py"]
    subprocess.check_call(cmd)
except subprocess.CalledProcessError as e:
    print(f"Error running dashboard: {e}")
    print("\nPress Enter to close...")
    input()
except KeyboardInterrupt:
    pass
