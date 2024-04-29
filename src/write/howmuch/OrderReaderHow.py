import re
from copy import deepcopy

from typing import List

from src.read.SpExReader import SpExReader
from src.read.model.Order import Order
from src.write.ecount.EcountFromSpExParser import EcountFromSpExParser

HOW_TITLE = ['순번', '거래일(예:2018-01-01)', '구분', '코드', '거래처명', '유형', '적요', '결제장부', '거래금액', '품목코드', '품목명', '규격', '단위', '수량',
             '단가', '공급가', '부가세', '합계금액', '창고코드', '창고명', '비고', '프로젝트(범 주)', '은행코드', '카드코드']


class OrderReaderHow:
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
        # HOW_TITLE = ['순번', '거래일(예:2018-01-01)', '구분', '코드', '거래처명', '유형', '적요', '결제장부', '거래금액', '품목코드', '품목명', '규격',
        #              '단위', '수량', '단가', '공급가', '부가세', '합계금액', '창고코드', '창고명', '비고', '프로젝트(범 주)', '은행코드', '카드코드']
        if key == '순번':
            return self.idx
        if key == '거래일(예:2018-01-01)':
            return self.order.identifier.date
        if key == '구분':
            return 1
        if key == '코드':
            return re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1]
        if key == '거래처명':
            return self.order.customer.nameAndCode
        if key == '유형':
            return 1  # TODO 부가세 0이면 2
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
        if key == '결제장부':
            return 4  # TODO 써비스면 5.어음으로 할까?
        if key == '거래금액':
            return self.order.totalPrice
        if key == '품목코드':
            return 1  # ?????????????????????
        if key == '품목명':
            return self.order.goods
        if key == '수량':
            return self.order.amount
        if key == '단가':
            return self.order.unitPrice
        if key == '공급가':
            return self.order.amount * self.order.unitPrice
        # if key == '부가세':
        #     return self.order.amount * self.order.vat
        #     return math.round(self.order.amount * self.order.unitPrice/10)
        if key == '합계금액':
            return self.order.amount * self.order.unitPrice

    def getDocumentsOfOrder(self) -> List[List]:
        if self.order.isFree:
            return [self.getDocByTitle(), self.getFreeDoc()]
        else:
            return [self.getDocByTitle()]

    def getDocByTitle(self) -> List:
        doc = []
        for key in ECOUNT_TITLE:
            doc.append(self[key])
        return doc

    def getFreeDoc(self) -> List:
        doc = []
        for key in ECOUNT_TITLE:
            if key in ['거래일(예:2018-01-01)', '구분', '코드', '거래처명', '유형', '적요', '결제장부', '거래금액']:
                doc.append(None)
                continue
            if key == '품목코드':
                doc.append('*')
                continue
            if key in ['수량', '단가', '공급가', '합계금액']:  # '부가세',
                doc.append(-self[key])
                continue
            doc.append(self[key])
        return doc


if __name__ == '__main__':
    spExReader = SpExReader()
    orders = spExReader.getOrders()
    eod = EcountFromSpExParser(orders[0], 1)
    print(eod['품목명'])
