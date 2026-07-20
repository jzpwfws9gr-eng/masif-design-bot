import os
import sys
import subprocess

if __name__ == "__main__":
    raise SystemExit(subprocess.call([sys.executable, "bot.py"], env=os.environ.copy()))
