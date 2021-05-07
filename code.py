import secrets
from tinydb import TinyDB, where


class Code:

    def __init__(self):
        db = TinyDB('codes.json')
        self.tb = db.table('Codes')

    @staticmethod
    def _code_gen():
        return secrets.token_urlsafe(5)

    def gen(self, amount: int):
        codes = []
        for i in range(amount):
            code = self._code_gen()
            self.tb.insert({
                'code': code
            })
            codes.append(code)
        return codes

    def check(self, code):
        return self.tb.get(where('code') == code)

    def del_code(self, code):
        self.tb.remove(
            where('code') == code
        )


if __name__ == '__main__':
    c = Code()
    c.gen(10)
