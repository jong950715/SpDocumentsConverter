import os
from tempfile import gettempdir


def getRootDir() -> str:
    root_dir = os.path.dirname(os.path.abspath(__file__))
    return root_dir


def getTempDir():
    return gettempdir().replace('\\\\', '/').replace('\\','/')


def main():
    r = getRootDir()
    t = getTempDir()
    print(r)
    print(t)


if __name__ == "__main__":
    main()
