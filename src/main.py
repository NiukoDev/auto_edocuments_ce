import os

os.environ.setdefault("TK_SILENCE_DEPRECATION", "1")

from ui import run_app


if __name__ == "__main__":
    run_app()
