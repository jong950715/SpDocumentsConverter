import re
from copy import deepcopy
from datetime import datetime

from typing import List

import openpyxl

from src.read.SpExReader import SpExReader
from src.read.model.Order import Order

WEHAGO_TITLE = ['년도', '월', '일', '매입매출구분(1 - 매출 / 2 - 매입)', '과세유형', '불공제사유', '신용카드거래처코드', '신용카드사명', '신용카드(가맹점)번호',
                '거래처명', '사업자(주민)등록번호', '공급가액', '부가세', '품명', '전자세금(1.전자)', '기본계정', '상대계정', '현금영수증 승인번호']


# SELF_USE_TITLE = ['일자', '순번', '거래처코드', '거래처명', '출하창고', '담당자', '적요(상단)', '품목코드', '품목명', '규격', '수량', '사용유형', '적요', ]


class WehagoFromSpExParser:

    def __init__(self, order, idx):
        self.order: Order = order
        self.idx: int = idx
        # self.parse()
        # re.compile('[#]([\d]+)')
        # outOrder[10] = order['customer']['idx'].zfill(6) + '-0000000'
        # outOrder[13] = '{0} {1} x \{2} {3}'.format(order['goods'], int(order['amount']), format(int(order['perPrice']), ','), order['customer']['name'])

    # def parse(self):
    # self.year, self.month, self.date = time.strftime('%Y', t), time.strftime('%m', t), time.strftime('%d', t)
    # self.customerCode = re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1]

    def __getitem__(self, key):
        item = self._getitem_(key)
        item = item.replace('\\', '￦') if isinstance(item, str) else item
        return item

    def _getitem_(self, key):
        # WEHAGO_TITLE = ['년도', '월', '일', '매입매출구분(1 - 매출 / 2 - 매입)', '과세유형', '불공제사유', '신용카드거래처코드', '신용카드사명',
        #                 '신용카드(가맹점)번호',
        #                 '거래처명', '사업자(주민)등록번호', '공급가액', '부가세', '품명', '전자세금(1.전자)', '기본계정', '상대계정', '현금영수증 승인번호']

        if key == '년도':
            return self.order.identifier.date.strftime('%Y')
        if key == '월':
            return self.order.identifier.date.strftime('%m')
        if key == '일':
            return self.order.identifier.date.strftime('%d')
        if key == '매입매출구분(1 - 매출 / 2 - 매입)':
            return 1
        if key == '과세유형':
            return 13
        if key == '사업자(주민)등록번호':
            return '{0}-0000000'.format(re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1].zfill(6))

        if key in ['공급가액', '부가세']:
            # from decimal import Decimal, getcontext, ROUND_HALF_UP
            # getcontext().rounding = ROUND_HALF_UP
            # return {'공급가액': int(round(Decimal(self.order.totalPrice) * 10 / 11, 0)) if self.order.vat else int(
            #     self.order.totalPrice),
            #         '부가세': int(round(Decimal(self.order.totalPrice) * 1 / 11, 0)) if self.order.vat else int(0),
            #         }[key]
            return {'공급가액': self.order.totalPrice, '부가세': None, }[key]
        if key in ['품명']:
            isSameSenderReceiver = (self.order.sender.name == self.order.receiver.name
                                    or self.order.sender.name is None
                                    or self.order.sender.name == '')
            # if self.order.sender.name in ['성풍물산', '수건어물']:
            return ('{0} {1} x \{2}'.format(self.order.goods,
                                            int(self.order.amount),
                                            format(int(self.order.unitPrice), ','),
                                            self.order.sender.name) +
                    (' {0}'.format(self.order.receiver.name) if isSameSenderReceiver else
                     ' {0} -> {1}'.format(self.order.sender.name, self.order.receiver.name))
                    )
        if key in ['기본계정', '상대계정']:
            return {
                '기본계정': 401,
                '상대계정': 108,
            }[key]

        # if key == '일자':
        #     return self.order.identifier.date.strftime('%Y/%m/%d')
        # if key == '순번':
        #     return self.idx
        # if key == '거래처코드':
        #     return 'C{0}'.format(re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1])
        # if key == '거래처명':
        #     return self.order.customer.nameAndCode
        # if key in ['이름', '주소', '전화번호', '이름(보내는 분)', '주소(보내는 분)', '전화번호(보내는 분)']:
        #     return {'이름': self.order.receiver.name,
        #             '주소': self.order.receiver.address,
        #             '전화번호': self.order.receiver.phone,
        #             '이름(보내는 분)': self.order.sender.name,
        #             '주소(보내는 분)': self.order.sender.address,
        #             '전화번호(보내는 분)': self.order.sender.phone,
        #             }[key]
        # if key == '거래유형':
        #     return 11 if self.order.vat else 12
        # if key in ['적요(상품)', '적요']:
        #     isSameSenderReceiver = (self.order.sender.name == self.order.receiver.name
        #                             or self.order.sender.name is None
        #                             or self.order.sender.name == '')
        #     # if self.order.sender.name in ['성풍물산', '수건어물']:
        #     return ('{0} {1} x \{2}'.format(self.order.goods,
        #                                     int(self.order.amount),
        #                                     format(int(self.order.unitPrice), ','),
        #                                     self.order.sender.name) +
        #             (' {0}'.format(self.order.receiver.name) if isSameSenderReceiver else
        #              ' {0} -> {1}'.format(self.order.sender.name, self.order.receiver.name))
        #             )
        # if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세']:
        #     from decimal import Decimal, getcontext, ROUND_HALF_UP
        #     getcontext().rounding = ROUND_HALF_UP
        #     return {'품목코드': 'G{0}'.format(self.order.goodsCode),
        #             '품목명': self.order.goods,
        #             '수량': int(self.order.amount),
        #             '단가': self.order.unitPrice,
        #             '공급가액': int(round(Decimal(self.order.totalPrice) * 10 / 11, 0)) if self.order.vat else int(self.order.totalPrice),
        #             '부가세': int(round(Decimal(self.order.totalPrice) * 1 / 11, 0)) if self.order.vat else int(0),
        #             }[key]

    def getDocumentsOfOrder(self) -> List[List]:
        if self.order.isFree:
            return [self.getFreeDocByTitle2()]
            # return [self.getDocByTitle(), self.getFreeDocByTitle()]
        else:
            return [self.getDocByTitle()]

    def getDocByTitle(self) -> List:
        doc = []
        for key in WEHAGO_TITLE:
            doc.append(self[key])
        return doc

    def getFreeDocByTitle2(self) -> List:
        # 임마는 one line으로 간다
        doc = []
        for key in WEHAGO_TITLE:
            if key in ['수량', '단가', '공급가액', '부가세', '거래유형']:
                doc.append({
                               '수량': self.order.amount,
                               '단가': 0,
                               '공급가액': 0,
                               '부가세': 0,
                               '거래유형': 10,
                           }[key]
                           )
                continue
            doc.append(self[key])
        return doc

    def getFreeDocByTitleSelfUse(self) -> List:
        # WEHAGO_TITLE = ['년도', '월', '일', '매입매출구분(1 - 매출 / 2 - 매입)', '과세유형', '불공제사유', '신용카드거래처코드', '신용카드사명',
        #                 '신용카드(가맹점)번호',
        #                 '거래처명', '사업자(주민)등록번호', '공급가액', '부가세', '품명', '전자세금(1.전자)', '기본계정', '상대계정', '현금영수증 승인번호']
        doc = []
        for key in WEHAGO_TITLE:
            if key in ['기본계정', '상대계정', '공급가액']:
                doc.append({
                               '기본계정': 411,
                               '상대계정': 411,
                               '공급가액': 1,
                           }[key]
                           )
                continue
            doc.append(self[key])
        return doc


if __name__ == '__main__':
    spExReader = SpExReader()
    orders = spExReader.getOrders()
    orw = WehagoFromSpExParser(orders[0], 1)
    print(orw['품목명'])
