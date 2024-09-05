import re
from copy import deepcopy, copy
from datetime import datetime

from typing import List, Dict, Union

import openpyxl
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from src.read.toggle.toggleLU import SOO_PACKAGE_FEE, ToggleLU
from src.read.toggle.ToggleReader import ToggleReader

SPEX_TITLES = ['담당자', '시간', '챙길것', '품목', '수량', '유형', '보험사', '지점', '주문자이름', '주소', '전화번호', '단가', '금액', '수금', ]


def copyCell(ws, input: Cell):
    output: Cell = Cell(ws, value=input.value)
    output.font = copy(input.font)
    output.fill = copy(input.fill)
    output.border = copy(input.border)
    output.alignment = copy(input.alignment)
    return output


def applyStylesAtValue(ws, styles, value):
    if value == '#NoStyle#':
        return Cell(ws, value='')
    output: Cell = Cell(ws, value=value)
    output.font = copy(styles['font'])
    output.fill = copy(styles['fill'])
    output.border = copy(styles['border'])
    output.alignment = copy(styles['alignment'])
    output.number_format = copy(styles['number_format'])
    return output


class SpExFromToggleParser:
    toggleLU = ToggleLU()
    styles = toggleLU.getStyles()
    sooLU = toggleLU.sooLU
    happyLU = toggleLU.happyLU

    def __init__(self, order, idx, ws: Union[None, Worksheet] = None):
        self.order: Dict = order
        self.idx: int = idx
        self.ws = ws
        self.isBundle = False

        if self.order['쇼핑몰'] == '수건어물':
            self._sooLU = self.sooLU[self.order]
        if self.order['쇼핑몰'] == '행복앤미소':
            self._happyLU = self.happyLU[self.order]

        self.shippingFee = 0
        # 수건어물은 룩업에서 가져오고
        # 저기 행복앤미소는 토글에서 설정해주고
        if self.order['쇼핑몰'] == '수건어물':
            self.goodsName = self._sooLU['출고지시서품목명'].value
            self.unitPrice = self._sooLU['납품단가'].value
            self.order['수량'] = self.order['수량'] * self._sooLU['납품수량'].value
        elif self.order['쇼핑몰'] == '행복앤미소':
            self.goodsName = self._happyLU['출고지시서품목명'].value
            self.unitPrice = self._happyLU['납품단가'].value
            self.order['수량'] = self.order['수량'] * self._happyLU['납품수량'].value
        else:
            self.goodsName = self.order['상품명'] + self.order['옵션정보']
            self.unitPrice = (self.order['상품별 총 주문금액'] - self.order['상품별 할인액']) / self.order['수량']

        goodsCodes = re.findall('[#]([\d]+)', self.goodsName)
        self.goodsCode = goodsCodes[-1] if len(goodsCodes) > 0 else 'G0'

        # self.parse()
        # re.compile('[#]([\d]+)')
        # outOrder[10] = order['customer']['idx'].zfill(6) + '-0000000'
        # outOrder[13] = '{0} {1} x \{2} {3}'.format(order['goods'], int(order['amount']), format(int(order['perPrice']), ','), order['customer']['name'])

    # def parse(self):
    # self.year, self.month, self.date = time.strftime('%Y', t), time.strftime('%m', t), time.strftime('%d', t)
    # self.customerCode = re.findall('[#]([\d]+)', self.order.customer.nameAndCode)[-1]

    def __getitem__(self, key) -> Union[int, float, str, Cell, None]:
        if self.isBundle and key in ['담당자']:
            return '#NoStyle#'
        if self.isBundle and key in ['담당자', '시간', '챙길것', '보험사', '지점', '주문자이름', '주소', '전화번호', '수금']:
            return ''
        item = self._getitem_(key)
        item = item.replace('\\', '￦') if isinstance(item, str) else item
        return item

    def _getitem_(self, key):
        # SPEX_TITLES = ['담당자', '시간', '챙길것', '품목', '수량', '보험사', '지점', '주문자이름', '주소', '전화번호', '단가', '금액', '수금', ]

        if key == '담당자':
            # return datetime.strptime(self.order['발송일'][0:10], '%Y.%M.%d').strftime('%Y-%M-%d')
            return '수' if self.order['쇼핑몰'] == '수건어물' else '수'
        if key == '시간':
            return datetime.strptime(self.order['발송일'][0:10], '%Y.%M.%d').strftime('%d.택배') if self.order[
                                                                                                   '발송일'] is not None and \
                                                                                               self.order['발송일'] != '' \
                else datetime.now().strftime('%d.택배')
        if key == '챙길것':
            return '수건어물' if self.order['쇼핑몰'] == '수건어물' else '행복앤미소'
        if key == '품목':
            return copyCell(self.ws, self._sooLU['출고지시서품목명']) if self.order['쇼핑몰'] == '수건어물' \
                else copyCell(self.ws, self._happyLU['출고지시서품목명']) if self.order['쇼핑몰'] == '행복앤미소' \
                else self.goodsName
        if key == '수량':
            return self.order['수량']
        if key == '유형':
            return self._sooLU['유형'].value if self.order['쇼핑몰'] == '수건어물' \
                else self._happyLU['유형'].value if self.order['쇼핑몰'] == '행복앤미소' \
                else ''
        if key == '보험사':
            return None
        if key == '지점':
            return None
        if key == '주문자이름':
            return self.order['수취인명']
        if key == '주소':
            return self.order['통합배송지']
        if key == '전화번호':
            return re.sub(r"(\d{3})(\d{3,4})(\d{4})", r"\1-\2-\3", self.order['수취인연락처1'])
        if key == '단가':
            return self.unitPrice
        if key == '금액':
            # return self.unitPrice * self.order['수량'] + self.shippingFee if self.order['쇼핑몰'] == '수건어물' else self.unitPrice * self.order['수량'] + self.order['배송비 합계']
            return "=L{0}*E{0}+{1}".format(self.ws.max_row+1, self.shippingFee)
        if key == '수금':
            return self.order['배송메세지']

    def getDocumentsOfOrder(self) -> List[List]:
        return [self.getDocByTitle()]
        # if self.order.isFree:
        #     return [self.getDocByTitle(), self.getFreeDocByTitle()]
        # else:
        #     return [self.getDocByTitle()]

    def getDocByTitle(self) -> List:
        doc = []
        for key in SPEX_TITLES:
            if isinstance(self[key], Cell):
                doc.append(self[key])
                continue
            doc.append(applyStylesAtValue(self.ws, self.styles[key], self[key]))
        return doc

    def setBundleOrder(self):
        self.isBundle = True

    def setShippingFee(self):
        self.shippingFee = self.order['배송비 합계'] + (SOO_PACKAGE_FEE if self.order['쇼핑몰'] == '수건어물' else 0)
        return self.shippingFee

    def getShippingDoc(self) -> List:
        doc = []
        for key in SPEX_TITLES:
            doc.append(self.getShippingEntry(key))
        return doc

    def getShippingEntry(self, key) -> Union[str, int]:
        if key in ['품목코드', '품목명', '수량', '단가', '공급가액', '부가세']:
            return {'품목코드': 'G택배'.format(self.goodsCode),
                    '품목명': '택배비',
                    '수량': 1,

                    '단가': self.order['배송비 합계'] + SOO_PACKAGE_FEE if self.order['쇼핑몰'] == '수건어물'
                    else self.order['배송비 합계'] if self.order['쇼핑몰'] == '행복앤미소' else None,

                    '공급가액': self.order['배송비 합계'] + SOO_PACKAGE_FEE if self.order['쇼핑몰'] == '수건어물'
                    else self.order['배송비 합계'] if self.order['쇼핑몰'] == '행복앤미소' else None,

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
