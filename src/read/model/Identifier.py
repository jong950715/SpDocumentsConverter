from typing import Dict


class Identifier:
    def __init__(self, identifier, date):
        self.identifier = identifier
        self.date = date

    @classmethod
    def fromRow(cls, row):
        ins = cls(None, None)
        if ins.verifyAndUpdateByRow(row):
            return ins
        else:
            raise Exception("Receiver not found", row)

    def toDict(self):
        return {
            'identifier': self.identifier,
            '날짜': self.date
        }

    def verifyRow(self, row: Dict[str, str]) -> bool:
        if _isIdentifier(row['챙길것']):
            return True
        return False

    def verifyAndUpdateByRow(self, row) -> bool:
        if self.verifyRow(row):
            self.identifier = row['챙길것']
            self.date = row['시간']
            return True
        return False


def _isIdentifier(s: str) -> bool:
    if not isinstance(s, str):
        return False
    for x in IDENTIFIERS:
        if s.startswith(x):
            return True
    else:
        return False


def getIdentifiers():
    return sorted(_getIdentifiers())


def _getIdentifiers():
    suffix = '#'
    ss1 = ['월', '화', '수', '목', '금', '토', '일']
    ss2 = ['요일', '욜']
    ss3 = ['택배', '배송']
    etc = ['행복앤미소', '수건어물']
    for s1 in ss1:
        for s2 in ss2:
            for s3 in ss3:
                yield '{}{}{}{}'.format(suffix, s1, s2, s3)

    for s in etc:
        yield '{}{}'.format(suffix, s)


IDENTIFIERS = getIdentifiers()

DUMMY_IDENTIFIER = Identifier(identifier='#더미더미', date='')
