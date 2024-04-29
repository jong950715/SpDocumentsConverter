import re
from copy import deepcopy
from datetime import datetime

from typing import List, Dict, Union

import openpyxl

from src.read.toggle.ToggleReader import ToggleReader

ECOUNT_TITLE = ['일자', '순번', '거래처코드', '거래처명', '담당자', '이름', '주소', '전화번호', '이름(보내는 분)', '주소(보내는 분)', '전화번호(보내는 분)', '거래유형',
                '적요', '품목코드', '품목명', '규격', '수량', '단가', '외화금액', '공급가액', '부가세', '적요(상품)', '생산전표생성', ]


class EcountFromToggleParser:
    def __init__(self, order, idx):
        self.order: Dict = order
        self.idx: int = idx
        self.unitPrice = (self.order['상품별 총 주문금액'] - self.order['상품별 할인액']) / self.order['수량']
        goodsCodes = re.findall('[#]([\d]+)', self.order['옵션정보'])
        self.goodsCode = goodsCodes[-1] if len(goodsCodes) > 0 else '0'
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
            return datetime.strptime(self.order['발송일'][0:10], '%Y.%M.%d').strftime('%Y-%M-%d')
        if key == '순번':
            return self.idx
        if key == '거래처코드':
            return 'C1691'
        if key == '거래처명':
            return '수건어물'
        if key in ['이름', '주소', '전화번호', '이름(보내는 분)', '주소(보내는 분)', '전화번호(보내는 분)']:
            return {'이름': self.order['수취인명'],
                    '주소': self.order['통합배송지'],
                    '전화번호': self.order['수취인연락처1'],
                    '이름(보내는 분)': self.order['구매자명'],
                    '주소(보내는 분)': '',
                    '전화번호(보내는 분)': self.order['구매자연락처'],
                    }[key]
        if key == '거래유형':
            return 11
        if key == '적요':
            isSameSenderReceiver = self.order['구매자명'] == self.order['수취인명']
            # if self.order.sender.name in ['성풍물산', '수건어물']:
            return ('{0} {1} x \{2}'.format(self.order['상품명'] + self.order['옵션정보'],
                                            self.order['수량'],
                                            format(int(self.unitPrice), ','),
                                            ) +
                    (' {0}'.format(self.order['수취인명']) if isSameSenderReceiver else
                     ' {0} -> {1}'.format(self.order['구매자명'], self.order['수취인명']))
                    )
        if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세']:
            return {'품목코드': 'G{0}'.format(self.goodsCode),
                    '품목명': self.order['상품명'] + self.order['옵션정보'],
                    '수량': self.order['수량'],
                    '단가': self.unitPrice,
                    '공급가액': self.unitPrice*self.order['수량'],
                    '부가세': 0,
                    }[key]

    def getDocumentsOfOrder(self) -> List[List]:
        return [self.getDocByTitle()]
        # if self.order.isFree:
        #     return [self.getDocByTitle(), self.getFreeDocByTitle()]
        # else:
        #     return [self.getDocByTitle()]

    def getDocByTitle(self) -> List:
        doc = []
        for key in ECOUNT_TITLE:
            doc.append(self[key])
        return doc

    def getShippingDoc(self) -> List:
        doc = []
        for key in ECOUNT_TITLE:
            doc.append(self.getShippingEntry(key))
        return doc

    def getShippingEntry(self, key) -> Union[str, int]:
        if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세']:
            return {'품목코드': 'G택배'.format(self.goodsCode),
                    '품목명': '택배비',
                    '수량': 1,
                    '단가': self.order['배송비 합계'],
                    '공급가액': self.order['배송비 합계'],
                    '부가세': 0,
                    }[key]
        return self[key]

    # def getFreeDocByTitle(self) -> List:
    #     doc = []
    #     for key in ECOUNT_TITLE:
    #         if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세']:
    #             doc.append(self.getFreeDocEntry(key))
    #             continue
    #         doc.append(self[key])
    #     return doc
    #
    # def getFreeDocEntry(self, key):
    #     if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세']:
    #         return {'품목코드': 'G0',
    #                 '품목명': '',
    #                 '수량': self.order.amount,
    #                 '단가': -self.order.unitPrice,
    #                 '공급가액': -self.order.totalPrice,
    #                 '부가세': 0,
    #                 }[key]
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
    toggleReader = ToggleReader.fromFileName('{0}/../TOGLE_출고내역서_수건어물_20240424_225939.xlsx'.format(getRootDir()))
    orders = toggleReader.getOrders()
    orw = EcountFromToggleParser(orders[0], 1)
    print(orw['품목명'])
