import re
from typing import Iterator


def isString(value) -> bool:
    if value is None or not isinstance(value, str) or value.isspace():
        return False
    return True


def test():
    # s = sorted(getIdentifiers())
    # print(type(s))
    # print(s)
    # print(isinstance(3,(int,float,str)))
    # print(isinstance(3.0, (int, float, str)))
    # print(isinstance('3.0', (int, float, str)))
    # print(isinstance('3.0', (int, float)))
    # print(isinstance(None, (int, float, str)))
    print(re.match('.*#\d+', '성풍#345'))
    print(re.match('.*#\d+', '#345'))
    print(re.match('.*#\d+', '##345'))
    print(re.match('.*#\d+', '345'))


if __name__ == '__main__':
    test()
