"""Operations that depend on the desktop operating system."""

import os
from pathlib import Path
import subprocess
import sys


def open_file(path):
    """Open a saved document with the user's default application."""
    path = str(Path(path).resolve())
    if sys.platform == 'win32':
        os.startfile(path)
    elif sys.platform == 'darwin':
        subprocess.run(['/usr/bin/open', path], check=True)
    else:
        raise RuntimeError('이 프로그램은 Windows와 macOS를 지원합니다.')
