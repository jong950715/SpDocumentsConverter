from src.read.tools import isString


class Sender:
    def __init__(self, name, address, phone):
        self.name = name
        self.address = address
        self.phone = phone

    def verifyAndUpdateByRow(self, row):
        if self.verifyRow(row):
            self.name = row['주문자이름']
            self.address = row['주소']
            self.phone = row['전화번호']
            return True
        return False

    def toDict(self):
        return {'보내는분-이름': self.name,
                '보내는분-주소': self.address,
                '보내는분-전화번호': self.phone}

    def verifyRow(self, row):
        if (isinstance(row['단가'], (int, float))
                and isinstance(row['금액'], (int, float))
                and isinstance(row['수량'], (int, float))):
            return False
        if '보내는분' not in row['품목'] and '보내는 분' not in row['품목']:
            return False
        if not isString(row['주문자이름']):
            return False
        return True


SENDER_SUNGPOONG = Sender(name="성풍물산", address='동호로', phone='010-0000-0000')
DUMMY_SENDER = SENDER_SUNGPOONG