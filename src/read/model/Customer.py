import re
from typing import Dict


class Customer:
    def __init__(self, customer, company=None):
        self.customer = customer
        self.company = company

    def verifyAndUpdateByRow(self, row):
        if self.verifyRow(row):
            self.customer = row['지점']
            self.company = row['보험사']
            return True
        return False

    def toDict(self):
        return {'거래처코드-지점': self.customer,
                '거래처코드-보험사': self.company}

    def verifyRow(self, row: Dict[str, str]) -> bool:
        return row['지점'] is not None and re.match('.*#\d+', row['지점']) is not None


DUMMY_CUSTOMER = Customer(customer='성풍물산#123', company='길음동지점')