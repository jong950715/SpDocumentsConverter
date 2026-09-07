from pathlib import Path
from tempfile import gettempdir


def getRootDir() -> str:
    return str(Path(__file__).resolve().parent)


def getTempDir():
    return gettempdir()


def main():
    r = getRootDir()
    t = getTempDir()
    print(r)
    print(t)


if __name__ == "__main__":
    main()
