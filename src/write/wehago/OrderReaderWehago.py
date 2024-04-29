import re

from typing import List

from src.definitions import getRootDir
from src.read.SpExReader import SpExReader
from src.read.model.Order import Order


class OrderReaderWehago:
    def __init__(self, order):
        self.order: Order = order
        # self.parse()
        # re.compile('[#]([\d]+)')
        # outOrder[10] = order['customer']['idx'].zfill(6) + '-0000000'
        # outOrder[13] = '{0} {1} x \{2} {3}'.format(order['goods'], int(order['amount']), format(int(order['perPrice']), ','), order['customer']['name'])

    # def parse(self):
        # self.year, self.month, self.date = time.strftime('%Y', t), time.strftime('%m', t), time.strftime('%d', t)
        # self.customerCode = re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1]

    def __getitem__(self, key):
        # ['년도', '월', '일', '매입매출구분(1 - 매출 / 2 - 매입)', '과세유형', '불공제사유', '신용카드거래처코드', '신용카드사명', '신용카드(가맹점)번호',
        # '거래처명', '사업자(주민)등록번호', '공급가액', '부가세', '품명', '전자세금(1.전자)', '기본계정', '상대계정', '현금영수증 승인번호']
        if key in ['년도', '월', '일']:
            return {'년도': self.order.identifier.date.year,
                    '월': self.order.identifier.date.month,
                    '일': self.order.identifier.date.day}[key]
        if key == '매입매출구분(1 - 매출 / 2 - 매입)':
            return 1
        if key == '과세유형':
            return 13
        if key == '거래처명':
            return self.order.customer.nameAndCode
        if key == '사업자(주민)등록번호':
            return re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1].zfill(6) + '-0000000'
        if key == '공급가액':
            return self.order.totalPrice
        if key == '품명':
            isSameSenderReceiver = self.order.sender.name == self.order.receiver.name
            # if self.order.sender.name in ['성풍물산', '수건어물']:
            return ('{0} {1} x \{2}'.format(self.order.goods,
                                               self.order.amount,
                                               format(int(self.order.unitPrice), ','),
                                               self.order.sender.name) +
                    (' {0}'.format(self.order.receiver.name) if isSameSenderReceiver else
                    ' {0} -> {1}'.format(self.order.sender.name, self.order.receiver.name))
                    )
        if key == '기본계정':
            return 401 if not self.order.isFree else 411
        if key == '상대계정':
            return 108 if not self.order.isFree else 411
        return None


    def getRowByTitle(self, title: List[str]):
        res = []
        for key in title:
            res.append(self[key])
        return res


if __name__ == '__main__':
    spExReader = SpExReader.fromFileName('{0}/../example.xlsx'.format(getRootDir()))
    orders = spExReader.getOrders()
    orw = OrderReaderWehago(orders[0])
    print(orw['품명'])
