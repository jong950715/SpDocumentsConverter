import copy

import openpyxl
import re

from typing import Dict

from src.read.model.Customer import DUMMY_CUSTOMER
from src.read.model.Identifier import Identifier, _isIdentifier, DUMMY_IDENTIFIER
from src.read.model.Order import Order
from src.read.model.Receiver import Receiver, DUMMY_RECEIVER
from src.read.model.Sender import SENDER_SUNGPOONG
from src.read.tools import isString

TITLES = ['시간', '챙길것', '품목', '수량', '보험사', '지점', '주문자이름', '주소', '전화번호', '단가', '금액', '수금']
MAX_COL = 13
MAX_ROW = 10


def onlyKCharacter(row):
    return list(map(lambda inStr: re.sub(r"[^가-힣]", "", str(inStr.value)), row))


def _isMemo(row: Dict[str, str]) -> bool:
    if (isinstance(row['단가'], (int, float))
            and isinstance(row['금액'], (int, float))
            and isinstance(row['수량'], (int, float))):
        return False
    return True


def _isOrder(row: Dict[str, str]) -> bool:
    if (not isinstance(row['단가'], int)
            or not isinstance(row['금액'], int)
            or not isinstance(row['수량'], int)):
        return False
    if not isString(row['품목']):
        return False
    return True


class SpExReader:
    def __init__(self):
        wb = openpyxl.load_workbook('../../example.xlsx', read_only=True, data_only=True)  # TODO : Definitions
        sheet = wb.active
        self.sheet = wb[sheet.title]
        self.titles, self.titleRow = self.getTitleRowInfo()
        self.identifier, self.sender, self.customer, self.receiver = DUMMY_IDENTIFIER, SENDER_SUNGPOONG, DUMMY_CUSTOMER, DUMMY_RECEIVER
        self.orders = []

    def test(self):
        orders = []
        for row in self.sheet.iter_rows(min_row=self.titleRow + 1, max_row=MAX_ROW, max_col=MAX_COL):
            execptionChecker = []  # for debug
            row = self.rowToDict(row)
            if self.identifier.verifyAndUpdateByRow(row):  # 구분자 솎아내기
                pass  # 기본 보내는 사람, 거래처 코드 설정
                continue
            if self.sender.verifyAndUpdateByRow(row):  # 보내는 사람
                execptionChecker.append('보내는 분')
            if self.customer.verifyAndUpdateByRow(row):  # 장부상 거래처코드
                execptionChecker.append('거래처코드')
            if _isMemo(row):
                continue
            if Order.verifyRow(row):
                self.addOrder(row)
                execptionChecker.append('주문')
            if len(execptionChecker) == 0:
                print("분기 안된 row가 있습니다!")
                print(row)
                pass  # 어떻게 처리 하실? 현재 파일 복사하고 거기에 에러메시지 출력 하면 될듯

        for o in self.orders:
            print(o.toDict())

    def addOrder(self, row):
        self.receiver.verifyAndUpdateByRow(row)
        o = Order(identifier=copy.deepcopy(self.identifier),
                  customer=copy.deepcopy(self.customer),
                  sender=copy.deepcopy(self.sender),
                  receiver=copy.deepcopy(self.receiver),
                  row=row)
        self.orders.append(o)

    def rowToDict(self, row) -> Dict[str, str]:
        res = {}
        for cell, title in zip(row, self.titles):
            res[title] = cell.value
        return res

    def getTitleRowInfo(self):
        for row in self.sheet.iter_rows(min_row=1, max_row=MAX_ROW, max_col=MAX_COL):
            if self.isTitleRow(row):
                return onlyKCharacter(row), row[0].row
        else:
            raise Exception("Title이 인식되지 않았습니다.")  # return false 하고 버그리포트

    def isTitleRow(self, row: tuple):
        row = onlyKCharacter(row)
        for title in TITLES:
            if title not in row:
                return False
        else:
            return True


if __name__ == '__main__':
    spExReader = SpExReader()
    spExReader.test()
