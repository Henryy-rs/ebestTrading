import unittest
import ebest

"""
id = your ebest id
passwd = password
cert_passwd = '공동인증서' password
"""

id = "abcd"
passwd = "1234"
cert_passwd = "5678"


class LoginTest(unittest.TestCase):
    def test_login_fail(self):
        session = ebest.InstXASession()
        self.assertEqual(0, session.login(id="abc", passwd="1234", cert_passwd="1234"))

    def test_login_success(self):
        session = ebest.InstXASession()
        self.assertEqual(1, session.login(id, passwd, cert_passwd))


class QueryTest(unittest.TestCase):
    def test_get_ticker_name_success(self):
        query = ebest.InstXAQueryT1102()
        self.assertEqual("삼성전자", query.get_ticker_symbol("005930"))

    def test_order_fail(self):
        query = ebest.InstXAQueryCSPAT00600()
        self.assertEqual(['주문실패', False], query.send_order(0, 0, "005930", 1, 70000, "신규매수", "지정가"))