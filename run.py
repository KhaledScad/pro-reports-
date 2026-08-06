"""Entry point: `python run.py` starts the server at http://127.0.0.1:8000 (local only).
To share it on the network again, change host to "0.0.0.0" below."""
import uvicorn

if __name__ == "__main__":
    uvicorn.run("app.app:app", host="127.0.0.1", port=8000, reload=False)
