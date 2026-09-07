import sys

from src.gui.MyGui import MyGui

def main():
    app = MyGui()
    if sys.argv[1:] == ['--smoke-test']:
        app.root.withdraw()
        app.root.update_idletasks()
        app.root.destroy()
    else:
        app.run()


if __name__ == '__main__':
    main()
