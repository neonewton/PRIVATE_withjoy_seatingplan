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


streamlit run "app.py"
