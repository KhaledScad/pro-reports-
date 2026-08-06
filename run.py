"""Entry point: `python run.py` starts the server at http://127.0.0.1:8000,
reachable from other devices on the same network at http://<this-machine-ip>:8000"""
import uvicorn

if __name__ == "__main__":
    uvicorn.run("app.app:app", host="0.0.0.0", port=8000, reload=False)
