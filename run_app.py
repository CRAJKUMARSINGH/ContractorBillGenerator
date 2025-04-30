import subprocess
import time
import os

def run_streamlit():
    # Start Streamlit server in a separate process
    subprocess.Popen(['streamlit', 'run', 'app.py', '--server.port', '8501'])
    
    # Wait a moment for the server to start
    time.sleep(2)
    
    # Get the URL
    url = 'http://localhost:8501'
    
    # Try to open Chrome
    try:
        chrome_path = None
        # Check common Chrome locations
        if os.path.exists(r'C:\Program Files\Google\Chrome\Application\chrome.exe'):
            chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
        elif os.path.exists(r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'):
            chrome_path = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
            
        if chrome_path:
            subprocess.Popen([chrome_path, url])
            return
            
    except Exception as e:
        print(f"Chrome not found: {e}")
    
    # If Chrome not found or failed, try Firefox
    try:
        firefox_path = None
        # Check common Firefox locations
        if os.path.exists(r'C:\Program Files\Mozilla Firefox\firefox.exe'):
            firefox_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
        elif os.path.exists(r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe'):
            firefox_path = r'C:\Program Files (x86)\Mozilla Firefox\firefox.exe'
            
        if firefox_path:
            subprocess.Popen([firefox_path, url])
            return
            
    except Exception as e:
        print(f"Firefox not found: {e}")
    
    # If both browsers fail, fall back to default browser
    import webbrowser
    webbrowser.open(url)

if __name__ == "__main__":
    run_streamlit()
