import win32com.client
import pythoncom


def ready_to_send_query():
    XAQueryEventHandler.query_state = 0


def ready_to_receive_query():
    while XAQueryEventHandler.query_state == 0:
        pythoncom.PumpWaitingMessages()


class XASessionEventHandler:
    login_state = 0

    def OnLogin(self, code, msg):
        if code == "0000":
            print("로그인 성공")
            XASessionEventHandler.login_state = 1
        else:
            print("로그인 실패")


class XAQueryEventHandler:
    query_state = 0

    def OnReceiveData(self, code):
        XAQueryEventHandler.query_state = 1


# 계정세션
class InstXASession:
    def __init__(self):
        self.session = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEventHandler)
        id = "" # 아이디
        passwd = "" # 비밀번호
        cert_passwd = ""    # 인증서 비밀번호
        self.session.ConnectServer("hts.ebestsec.co.kr", 20001)
        self.session.Login(id, passwd, cert_passwd, 0, 0)

        while XASessionEventHandler.login_state == 0:
            pythoncom.PumpWaitingMessages()

    def get_account_code(self):
        account_list = []
        num_account = self.session.GetAccountListCount()
        for i in range(num_account):
            account = self.session.GetAccountList(i)
            account_list.append(account)
        return account_list


class InstXAQuery:
    def __init_subclass__(cls):
        super().__init_subclass__()
        cls.query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEventHandler)
        cls.tr_code = cls.__name__[11:].lower()   # 전송코드
        cls.query.ResFileName = "C:\\eBEST\\xingAPI\\Res\\%s.res" % cls.tr_code     # Res
        cls.in_block_list = []
        cls.out_block_list = []

    def make_blocks(self, num_in_block=0, num_out_block=0):     # 전송블록, 수신블록
        if num_in_block == 0 and num_out_block == 0:
            self.in_block_list = ["%sInBlock" % self.tr_code].copy()
            self.out_block_list = ["%sOutBlock" % self.tr_code].copy()
        elif num_in_block == 0:
            self.in_block_list = ["%sInBlock" % self.tr_code].copy()
            self.out_block_list = ["%sOutBlock" % self.tr_code] + \
                                  list(map(lambda x: "%sOUTBlock" % self.tr_code + str(x+1), range(num_out_block)))
        else:
            self.in_block_list = list(map(lambda x: "%sInBlock" % self.tr_code + str(x+1), range(num_in_block))).copy()
            self.out_block_list = list(map(lambda x: "%sOUTBlock" % self.tr_code + str(x+1), range(num_out_block))).copy()


# 단일조회
class InstXAQueryT1102(InstXAQuery):
    def __init__(self):
        self.make_blocks()

    def get_name_of_ticker(self, ticker_code):
        if len(ticker_code) != 6:
            return "유효하지 않은 코드"

        ready_to_send_query()
        self.query.SetFieldData(self.in_block_list[0], "shcode", 0, ticker_code)
        self.query.Request(False)

        ready_to_receive_query()

        received_message = self.query.GetFieldData(self.out_block_list[0], "hname", 0)
        if received_message == "":
            return "유효하지 않은 코드"
        else:
            return received_message

    def get_price_of_ticker(self, ticker_code):
        if len(ticker_code) != 6:
            return 0

        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "shcode", 0, ticker_code)
        self.query.Request(False)

        ready_to_receive_query()

        received_message = self.query.GetFieldData(self.out_block_list[0], "price", 0)

        if received_message == "":
            return 0
        else:
            return int(received_message)


# 주문
class InstXAQueryCSPAT00600(InstXAQuery):
    def __init__(self):
        self.make_blocks(1, 2)

    def send_order(self, account_number, account_pwd, ticker_code, amount, price, order_type, price_type="지정가"):
        if order_type == "신규매수":
            sell_or_buy = 2
        elif order_type == "신규매도":
            sell_or_buy = 1

        if price_type == "지정가":
            manual_or_normal_price = "00"
        elif price_type == "시장가":
            manual_or_normal_price = "03"

        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "AcntNo", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "InptPwd", 0, account_pwd)
        self.query.SetFieldData(self.in_block_list[0], "IsuNo", 0, ticker_code)
        self.query.SetFieldData(self.in_block_list[0], "OrdQty", 0, amount)
        self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, price)
        self.query.SetFieldData(self.in_block_list[0], "BnsTpCode", 0, sell_or_buy)    # 매도(1), 매수(2)
        self.query.SetFieldData(self.in_block_list[0], "OrdprcPtnCode", 0, manual_or_normal_price)  # 호가유형코드
        self.query.SetFieldData(self.in_block_list[0], "MgntrnCode", 0, '000')  # 신용거래코드
        self.query.SetFieldData(self.in_block_list[0], "LoanDt", 0, '0')  # 대출일
        self.query.SetFieldData(self.in_block_list[0], "OrdCndiTpCode", 0, '0')  # 주문조건구분
        self.query.Request(False)

        ready_to_receive_query()

        order_list = []
        num_blocks = self.query.GetBlockCount(self.out_block_list[0])
        for i in range(num_blocks):
            order_list.append([])
            order_list.append[i](self.query.GetFieldData(self.out_block_list[1], "OrdNo", i))   # 주문번호
            order_list.append[i](self.query.GetFieldData(self.out_block_list[0], "OrdPrc", i))  # 주문가
            order_list.append[i](self.query.GetFieldData(self.out_block_list[0], "BnsTpCode", i))  # 매매구분

        if num_blocks == 0:
            order_list = ["error: order failed"]

        print(order_list)
        return order_list


class InstXAQueryT0424(InstXAQuery):
    def __init__(self):
        self.make_blocks(0, 1)

    def get_account_balance(self, account_number, account_pwd):
        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "accno", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "passwd", 0, account_pwd)
        self.query.SetFieldData(self.in_block_list[0], "prcgb", 0, 1)   # 단가구분: 평균단가(1), BEP단가(2)
        self.query.SetFieldData(self.in_block_list[0], "chegb", 0, 0)   # 체결구분: 결제기준잔고(0), 체결기준잔고(2)
        self.query.SetFieldData(self.in_block_list[0], "dangb", 0, 0)   # 단일가구분: 정규장(0), 시장외단일가(1)
        self.query.SetFieldData(self.in_block_list[0], "charge", 0, 0)  # 제비용포함여부: 미포함(0), 포함(1)
        self.query.Request(False)

        ready_to_receive_query()

        account_info_list = []
        account_info_list.append(self.query.GetFieldData(self.out_block_list[0], "sunamt", 0))
        account_info_list.append(self.query.GetFieldData(self.out_block_list[0], "dtsunik", 0))
        account_info_list.append(self.query.GetFieldData(self.out_block_list[0], "mamt", 0))
        account_info_list.append(self.query.GetFieldData(self.out_block_list[0], "sunamt1", 0))
        account_info_list.append(self.query.GetFieldData(self.out_block_list[0], "tappamt", 0))
        account_info_list.append(self.query.GetFieldData(self.out_block_list[0], "tdtsunik", 0))
        account_info_list.append(str(round((int(account_info_list[0])/10000-1)*100, 3))+"%")    # 최초투자금 10000원

        return account_info_list

    def get_holding_stocks(self, account_number, account_pwd):
        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "accno", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "passwd", 0, account_pwd)
        self.query.SetFieldData(self.in_block_list[0], "prcgb", 0, 1)  # 단가구분: 평균단가(1), BEP단가(2)
        self.query.SetFieldData(self.in_block_list[0], "chegb", 0, 0)  # 체결구분: 결제기준잔고(0), 체결기준잔고(2)
        self.query.SetFieldData(self.in_block_list[0], "dangb", 0, 0)  # 단일가구분: 정규장(0), 시장외단일가(1)
        self.query.SetFieldData(self.in_block_list[0], "charge", 0, 0)  # 제비용포함여부: 미포함(0), 포함(1)
        self.query.SetFieldData(self.in_block_list[0], "cts_expcode", 0, "")    # 연속 조회시 사용
        self.query.Request(False)

        ready_to_receive_query()

        holding_stocks_info_list = []
        num_blocks = self.query.GetBlockCount(self.out_block_list[1])
        for i in range(num_blocks):
            holding_stocks_info_list.append([])
            holding_stocks_info_list[i].append(self.query.GetFieldData(self.out_block_list[1], "hname", i))
            holding_stocks_info_list[i].append(self.query.GetFieldData(self.out_block_list[1], "expcode", i))
            holding_stocks_info_list[i].append(self.query.GetFieldData(self.out_block_list[1], "janqty", i))
            holding_stocks_info_list[i].append(self.query.GetFieldData(self.out_block_list[1], "mamt", i))
            holding_stocks_info_list[i].append(self.query.GetFieldData(self.out_block_list[1], "price", i))
            holding_stocks_info_list[i].append(self.query.GetFieldData(self.out_block_list[1], "dtsunik", i))
            holding_stocks_info_list[i].append(self.query.GetFieldData(self.out_block_list[1], "sunikrt", i))

        return holding_stocks_info_list