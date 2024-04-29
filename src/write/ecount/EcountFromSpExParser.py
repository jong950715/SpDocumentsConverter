import re
from copy import deepcopy

from typing import List

import openpyxl

from src.read.SpExReader import SpExReader
from src.read.model.Order import Order

ECOUNT_TITLE = ['일자', '순번', '거래처코드', '거래처명', '담당자', '이름', '주소', '전화번호', '이름(보내는 분)', '주소(보내는 분)', '전화번호(보내는 분)', '거래유형',
                '적요', '품목코드', '품목명', '규격', '수량', '단가', '외화금액', '공급가액', '부가세', '적요(상품)', '생산전표생성', ]


class EcountFromSpExParser:
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
        # ECOUNT_TITLE = ['일자', '순번', '거래처코드', '거래처명', '담당자', '이름', '주소', '전화번호', '이름(보내는 분)', '주소(보내는 분)', '전화번호(보내는 분)',
        #                 '거래유형', '적요', '품목코드', '품목명', '규격', '수량', '단가', '외화금액', '공급가액', '부가세', '적요', '생산전표생성', ]

        if key == '일자':
            return self.order.identifier.date.strftime('%Y/%m/%d')
        if key == '순번':
            return self.idx
        if key == '거래처코드':
            return 'C{0}'.format(re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1])
        if key == '거래처명':
            return self.order.customer.nameAndCode
        if key in ['이름', '주소', '전화번호', '이름(보내는 분)', '주소(보내는 분)', '전화번호(보내는 분)']:
            return {'이름': self.order.receiver.name,
                    '주소': self.order.receiver.address,
                    '전화번호': self.order.receiver.phone,
                    '이름(보내는 분)': self.order.sender.name,
                    '주소(보내는 분)': self.order.sender.address,
                    '전화번호(보내는 분)': self.order.sender.phone,
                    }[key]
        if key == '거래유형':
            return 12
        if key == '적요':
            isSameSenderReceiver = self.order.sender.name == self.order.receiver.name
            # if self.order.sender.name in ['성풍물산', '수건어물']:
            return ('{0} {1} x \{2}'.format(self.order.goods,
                                            self.order.amount,
                                            format(int(self.order.unitPrice), ','),
                                            self.order.sender.name) +
                    (' {0}'.format(self.order.receiver.name) if isSameSenderReceiver else
                     ' {0} -> {1}'.format(self.order.sender.name, self.order.receiver.name))
                    )
        if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세']:
            from decimal import Decimal, getcontext, ROUND_HALF_UP
            getcontext().rounding = ROUND_HALF_UP
            return {'품목코드': 'G{0}'.format(self.order.goodsCode),
                    '품목명': self.order.goods,
                    '수량': self.order.amount,
                    '단가': self.order.unitPrice,
                    '공급가액': round(Decimal(self.order.totalPrice) * 10 / 11, 0),
                    '부가세': round(Decimal(self.order.totalPrice) * 1 / 11, 0),
                    }[key]

    def getDocumentsOfOrder(self) -> List[List]:
        if self.order.isFree:
            return [self.getFreeDocByTitle2()]
            # return [self.getDocByTitle(), self.getFreeDocByTitle()]
        else:
            return [self.getDocByTitle()]

    def getDocByTitle(self) -> List:
        doc = []
        for key in ECOUNT_TITLE:
            doc.append(self[key])
        return doc

    def getFreeDocByTitle2(self) -> List:
        # 임마는 one line으로 간다
        doc = []
        for key in ECOUNT_TITLE:
            if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세', '거래유형']:
                doc.append({'품목코드': 'G샘플',
                            '품목명': '',
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

    def getFreeDocByTitle(self) -> List:
        doc = []
        for key in ECOUNT_TITLE:
            if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세', '거래유형']:
                doc.append(self.getFreeDocEntry(key))
                continue
            doc.append(self[key])
        return doc

    def getFreeDocEntry(self, key):
        if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세', '거래유형']:
            return {'품목코드': 'G샘플',
                    '품목명': '',
                    '수량': self.order.amount,
                    '단가': -self.order.unitPrice,
                    '공급가액': -self.order.totalPrice,
                    '부가세': 0,
                    '거래유형': 10,
                    }[key]
    # def getFreeDoc(self) -> List:
    #     doc = []
    #     for key in ECOUNT_TITLE:
    #         if key in ['거래일(예:2018-01-01)', '구분', '코드', '거래처명', '유형', '적요', '결제장부', '거래금액']:
    #             doc.append(None)
    #             continue
    #         if key == '품목코드':
    #             doc.append('*')
    #             continue
    #         if key in ['수량', '단가', '공급가', '합계금액']:  # '부가세',
    #             doc.append(-self[key])
    #             continue
    #         doc.append(self[key])
    #     return doc


if __name__ == '__main__':
    spExReader = SpExReader()
    orders = spExReader.getOrders()
    orw = EcountFromSpExParser(orders[0], 1)
    print(orw['품목명'])
