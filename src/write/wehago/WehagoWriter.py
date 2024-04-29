import time

import openpyxl

from src.definitions import getRootDir
from src.read.SpExReader import SpExReader
from src.write.wehago.OrderReaderWehago import OrderReaderWehago

WEHAGO_TITLE = ['년도', '월', '일', '매입매출구분(1 - 매출 / 2 - 매입)', '과세유형', '불공제사유', '신용카드거래처코드', '신용카드사명', '신용카드(가맹점)번호',
                '거래처명', '사업자(주민)등록번호', '공급가액', '부가세', '품명', '전자세금(1.전자)', '기본계정', '상대계정', '현금영수증 승인번호']


# 품목	수량	단가	금액	identifier	날짜	거래처코드-지점	거래처코드-보험사	보내는분-이름	보내는분-주소	보내는분-전화번호	받는분-이름	받는분-주소	받는분-전화번호
class WehagoWriter:
    def __init__(self, sheet):
        self.sheet = sheet

    def getDocsFromSpEx(self):
        # spExReader = SpExReader('{0}/../example.xlsx'.format(getRootDir()))
        spExReader = SpExReader.fromSheet(self.sheet)
        orders = spExReader.getOrders()
        wb = openpyxl.Workbook()
        ws = wb.active
        wehagoDocuments = ws
        wehagoDocuments.append(WEHAGO_TITLE)
        for o in orders:
            document = OrderReaderWehago(o).getRowByTitle(WEHAGO_TITLE)
            wehagoDocuments.append(document)

        wb.save('{0}/{1}-{2}.xlsx'.format(getRootDir(), 'WehagoWriter', time.strftime("%Y%m%d-%H%M")))


if __name__ == '__main__':
    w = WehagoWriter()
    w.getDocsFromSpEx()
