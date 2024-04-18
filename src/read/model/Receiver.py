import copy

from src.read.tools import isString


class Receiver:
    def __init__(self, name, address, phone):
        self.name = name
        self.address = address
        self.phone = phone

    @classmethod
    def fromRow(cls, row):
        ins = cls(None, None, None)
        if ins.verifyAndUpdateByRow(row):
            return ins
        else:
            raise Exception("Receiver not found", row)

    def toDict(self):
        return {'받는분-이름': self.name,
                '받는분-주소': self.address,
                '받는분-전화번호': self.phone}

    def verifyAndUpdateByRow(self, row) -> bool:
        if (not isString(row['주문자이름']) or
                not isString(row['주소']) or
                not isString(row['전화번호'])):
            return False
        self.name = row['주문자이름']
        self.address = row['주소']
        self.phone = row['전화번호']
        return True


DUMMY_RECEIVER = Receiver.fromRow({'주문자이름': '홍길동',
                                   '주소': '서울특별시',
                                   '전화번호': '01000000000'})

if __name__ == '__main__':
    print(DUMMY_RECEIVER.toDict())
    d = copy.deepcopy(DUMMY_RECEIVER)
    DUMMY_RECEIVER.phone = '029111'
    print(d.toDict())
    print(DUMMY_RECEIVER.toDict())
