# Create virtual environment (replace `venv` with your preferred folder)
python3 -m venv venv

# Activate (macOS/Linux)
source venv/bin/activate

# Activate (Windows)
venv\Scripts\activate

# Check Python Version
python3 -V

# Install Required Packages
pip install -r requirements.txt

# Python Interpreter
Make sure your script uses the correct Python interpreter path:
#!/usr/local/bin/python3
In VS Code, use ⇧ + ⌘ + P → “Python: Select Interpreter” and choose Python 3.13 manually.

⚠️ Avoid using:
#!/usr/bin/env python3


"""
# Ensure you have started Chrome with:
r"""
MAC:
open -a "Google Chrome" --args \
  --remote-debugging-port=9222 \
  --user-data-dir="$HOME/chrome-debug-profile"
then
curl http://127.0.0.1:9222/json
WINDOWS:
& 'C:\Program Files\Google\Chrome\Application\chrome.exe' `
  --remote-debugging-port=9222 `
  --user-data-dir="C:\Users\Neone\chrome-debug-profile"
then
curl http://127.0.0.1:9222/json

"""


# streamlit run "1_Elentra_iLAMS_atm_tool_V6/home.py"
