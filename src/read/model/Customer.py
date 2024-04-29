import re
from typing import Dict


class Customer:
    def __init__(self, nameAndCode, company=None):
        self.nameAndCode = nameAndCode
        self.company = company

    def verifyAndUpdateByRow(self, row):
        if self.verifyRow(row):
            self.nameAndCode = row['지점']
            self.company = row['보험사']
            return True
        return False

    def toDict(self):
        return {'거래처코드-지점': self.nameAndCode,
                '거래처코드-보험사': self.company}

    def verifyRow(self, row: Dict[str, str]) -> bool:
        return isinstance(row['지점'], str) and re.match('.*#\d+', row['지점']) is not None

    @classmethod
    def getByIdentifier(cls, identifier: str):
        if "배송" in identifier:
            return Customer(
                nameAndCode="성풍물산#0000",
            )

        # if 택배
        if "수건어물" in identifier:
            return Customer(
                nameAndCode="수건어물#1691"
            )
        elif "행복앤미소" in identifier:
            return Customer(
                nameAndCode="행복앤미소#5002"
            )
        elif "택배" in identifier:
            return Customer(
                nameAndCode="성풍물산#0000",
            )
        
        raise Exception(identifier)


DUMMY_CUSTOMER = Customer(nameAndCode='성풍물산@#123', company='길음동지점')
