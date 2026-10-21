# Start the web app (http://127.0.0.1:8000)
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$root\.venv\Scripts\python.exe" -m uvicorn backend.app:app --host 127.0.0.1 --port 8000 --app-dir $root
