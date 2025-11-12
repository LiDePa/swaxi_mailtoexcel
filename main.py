import asyncio
import os

from gui import app


def _configure_event_loop() -> None:
    """Ensure a compatible event loop policy on Windows environments."""
    if os.name == "nt":
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        except AttributeError:
            # Fallback for Python builds without WindowsSelectorEventLoopPolicy
            pass


if __name__ == "__main__":
    _configure_event_loop()
    debug_enabled = os.getenv("FLASK_DEBUG", "0") == "1"
    host = os.getenv("FLASK_HOST", "127.0.0.1")
    port = int(os.getenv("FLASK_PORT", "5000"))

    # Open the default web browser to the Flask app URL
    import threading
    import webbrowser

    def open_browser():
        url = f"http://{host}:{port}/"
        webbrowser.open(url)

    threading.Timer(1.0, open_browser).start()
    app.run(host=host, port=port, debug=debug_enabled)


        





        