import re
from typing import Dict, List

from src.read.model.Customer import DUMMY_CUSTOMER
from src.read.model.Identifier import DUMMY_IDENTIFIER
from src.read.model.Receiver import DUMMY_RECEIVER
from src.read.model.Sender import DUMMY_SENDER
from src.read.tools import isString

FREE_WORDS = ['#샘플', '#쌤플', '#서비스', '#써비스']


def doesWordContainWords(inputWord, words: List[str]):
    if not isinstance(inputWord, str):
        return False
    for searchWord in words:
        if searchWord in inputWord:
            return True
    else:
        return False


class Order:
    def __init__(self, identifier, customer, sender, receiver, row):
        self.identifier = identifier
        self.customer = customer
        self.sender = sender
        self.receiver = receiver
        # self.orderDate = row['']
        self.goods = row['품목']
        goodsCodes = re.findall('[#]([\d]+)', row['품목'])
        self.goodsCode = goodsCodes[-1] if len(goodsCodes) > 0 else '0'
        self.amount = row['수량']
        self.unitPrice = row['단가']
        self.totalPrice = row['금액']
        self.isFree = (doesWordContainWords(row['챙길것'], FREE_WORDS)
                       or doesWordContainWords(row['보험사'], FREE_WORDS))
        # self.logisticPrice = row['배송비??']
        self.vat = self.isVat(row['부가세'])
        # self.vat = True

    def toDict(self):
        l = [self.identifier, self.customer, self.sender, self.receiver]
        res = self._toDict()
        for x in l:
            res.update(x.toDict())
        return res

    def isVat(self, value):
        if value is None or not isinstance(value, str):
            return True  # 일단 면세로
            # return False  # 일단 면세로
        if value in ['O', 'o', '과세']:
            return True
        elif value in ['X', 'x', '면세']:
            return False
        else:
            pass  # TODO: 에러처리!

    def _toDict(self):
        return {
            '품목': self.goods,
            '품목코드': self.goodsCode,
            '수량': self.amount,
            '단가': self.unitPrice,
            '금액': self.totalPrice,
            '무상여부': self.isFree
        }

    @staticmethod
    def verifyRow(row: Dict[str, str]) -> bool:
        if (not isinstance(row['단가'], int)
                or not isinstance(row['금액'], int)
                or not isinstance(row['수량'], int)):
            return False
        if not isString(row['품목']):
            return False
        return True


DUMMY_ROW = {'우': '향', '시간': '17.중', '챙길것': '보자기XXX', '품목': '대추방울토마토3kg(4팩) ', '수량': 1, '보험사': '#부록,명세完',
             '지점': '명륜지점#409', '주문자이름': 'dummy', '주소': '서울시 동대문구 천호대로 429(장안동)  대성타워 3층 <한화 명륜지점>',
             '전화번호': '010-3156-4537', '단가': 35000, '금액': 35000, '수금': None, '부가세': None}
DUMMY_ORDER = Order(identifier=DUMMY_IDENTIFIER, customer=DUMMY_CUSTOMER, sender=DUMMY_SENDER, receiver=DUMMY_RECEIVER,
                    row=DUMMY_ROW)
if __name__ == '__main__':
    print(DUMMY_ORDER.toDict())
