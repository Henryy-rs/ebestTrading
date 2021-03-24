import win32com.client
import pythoncom
import pandas as pd
import datetime


investment = 0


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
        id = ""
        passwd = ""
        cert_passwd = ""
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

    def make_blocks(self, num_in_block=0, num_out_block=0):     # 전송블록, 수신블록 생성
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

    def get_name_of_ticker(self, stock_ticker):
        if len(stock_ticker) != 6:
            return "유효하지 않은 코드"

        ready_to_send_query()
        self.query.SetFieldData(self.in_block_list[0], "shcode", 0, stock_ticker)
        self.query.Request(False)

        ready_to_receive_query()

        received_message = self.query.GetFieldData(self.out_block_list[0], "hname", 0)
        if received_message == "":
            return "유효하지 않은 코드"
        else:
            return received_message

    def get_price_of_ticker(self, stock_ticker):
        if len(stock_ticker) != 6:
            return 0

        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "shcode", 0, stock_ticker)
        self.query.Request(False)

        ready_to_receive_query()

        received_message = self.query.GetFieldData(self.out_block_list[0], "price", 0)

        if received_message == "":
            return 0
        else:
            return int(received_message)

    def get_market_cap(self, stock_ticker):
        if len(stock_ticker) != 6:
            return 0

        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "shcode", 0, stock_ticker)
        self.query.Request(False)

        ready_to_receive_query()

        received_message = self.query.GetFieldData(self.out_block_list[0], "total", 0)

        if received_message == "":
            return 0
        else:
            return int(received_message)


# 주문
class InstXAQueryCSPAT00600(InstXAQuery):
    def __init__(self):
        self.make_blocks(1, 2)

    def send_order(self, account_number, account_pwd, stock_ticker, order_quantity, order_price, order_type, price_type="지정가", order_option=0):
        if order_type == "신규매수" or order_type == "종가매수" or order_type == "단일가매수":
            sell_or_buy = 2
        elif order_type == "신규매도" or order_type == "종가매도" or order_type == "단일가매도":
            sell_or_buy = 1

        ready_to_send_query()

        if price_type == "지정가":
            quote_type = "00"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, order_price)
        elif price_type == "시장가":
            quote_type = "03"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, 0)
        elif price_type == "종가":
            now = datetime.datetime.now()
            if 8 < now.hour < 9:
                quote_type = "61"
            else:
                quote_type = "81"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, 0)
        elif price_type == "단일가":
            quote_type = "82"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, order_price)
        elif price_type == "최유리지정가":
            quote_type = "06"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, 0)
        self.query.SetFieldData(self.in_block_list[0], "AcntNo", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "InptPwd", 0, account_pwd)
        self.query.SetFieldData(self.in_block_list[0], "IsuNo", 0, stock_ticker)
        self.query.SetFieldData(self.in_block_list[0], "OrdQty", 0, order_quantity)
        self.query.SetFieldData(self.in_block_list[0], "BnsTpCode", 0, sell_or_buy)    # 매도(1), 매수(2)
        self.query.SetFieldData(self.in_block_list[0], "OrdprcPtnCode", 0, quote_type)  # 호가유형코드
        self.query.SetFieldData(self.in_block_list[0], "MgntrnCode", 0, '000')  # 신용거래코드
        self.query.SetFieldData(self.in_block_list[0], "LoanDt", 0, '0')  # 대출일
        self.query.SetFieldData(self.in_block_list[0], "OrdCndiTpCode", 0, order_option)  # 주문조건구분
        self.query.Request(False)

        ready_to_receive_query()

        order_result_list = []
        num_blocks = self.query.GetBlockCount(self.out_block_list[0])
        print(num_blocks)
        #for i in range(num_blocks):
            #order_list.append([])
        order_result_list.append(self.query.GetFieldData(self.out_block_list[1], "OrdNo", 0))   # 주문번호
        order_time = self.query.GetFieldData(self.out_block_list[1], "OrdTime", 0)
        order_time = order_time[0:2] + ":" + order_time[2:4] + ":" + order_time[4:6]
        order_result_list.append(order_time)                                                    # 주문시각
        order_result_list.append(self.query.GetFieldData(self.out_block_list[0], "IsuNo", 0))   # 종목번호
        order_result_list.append(self.query.GetFieldData(self.out_block_list[0], "IsuNm", 0))   # 종목명
        order_result_list.append(order_type[2:])                                                # 주문종류
        order_result_list.append(self.query.GetFieldData(self.out_block_list[0], "OrdPrc", 0))  # 주문가
        order_result_list.append(self.query.GetFieldData(self.out_block_list[0], "OrdQty", 0))  # 주문개수

        if num_blocks == 0:
            return ["주문실패", False]  # 주문 실패시 모든 주문 공통적으로 이 메세지 반환
        return order_result_list


class InstXAQueryT0424(InstXAQuery):
    def __init__(self):
        self.make_blocks(0, 1)

    def get_account_balance_info(self, account_number, account_pwd=""):
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
        sunamt = self.query.GetFieldData(self.out_block_list[0], "sunamt", 0)       # 추정자산
        dtsunik = self.query.GetFieldData(self.out_block_list[0], "dtsunik", 0)     # 실현손익(하루동안)
        sunamt1 = self.query.GetFieldData(self.out_block_list[0], "sunamt1", 0)     # d2예수금
        mamt = self.query.GetFieldData(self.out_block_list[0], "mamt", 0)           # 매입금액
        tappamt = self.query.GetFieldData(self.out_block_list[0], "tappamt", 0)     # 평가금액
        tdtsunik = self.query.GetFieldData(self.out_block_list[0], "tdtsunik", 0)   # 평가손익
        
        try:
            prftrt = str(round((int(sunamt)/investment-1)*100, 3))+"%"   # 손익률(investment: 투자금)
        except ZeroDivisionError and ValueError:
            prftrt = "----"
        account_info = {"estimated_amount": sunamt, "realized_profit": dtsunik, "deposit_received": sunamt1,
                        "pchs_amount": mamt, "eval_amount": tappamt, "eval_profit": tdtsunik, "profit_rate": prftrt}
        account_info_list.append(account_info)
        return account_info_list

    def get_holding_stocks_info(self, account_number, account_pwd=""):
        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "accno", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "passwd", 0, account_pwd)   # 비밀번호 없이 조회가능
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

            hname = self.query.GetFieldData(self.out_block_list[1], "hname", i)
            expcode = self.query.GetFieldData(self.out_block_list[1], "expcode", i)
            janqty = self.query.GetFieldData(self.out_block_list[1], "janqty", i)
            if int(janqty) == 0:
                continue
            mamt = self.query.GetFieldData(self.out_block_list[1], "mamt", i)
            try:
                avrgprice = str(int(int(mamt)/int(janqty)))
            except ZeroDivisionError:
                avrgprice = "0"
            price = self.query.GetFieldData(self.out_block_list[1], "price", i)
            dtsunik = self.query.GetFieldData(self.out_block_list[1], "dtsunik", i)
            sunikrt = self.query.GetFieldData(self.out_block_list[1], "sunikrt", i)
            stock_info = {"stock_name": hname, "stock_ticker": expcode, "quantity": janqty, "pchs_amount": mamt,
                          "avrg_price": avrgprice, "cur_price": price, "profit": dtsunik, "profit_rate": sunikrt}
            holding_stocks_info_list.append(stock_info)
        return holding_stocks_info_list


class InstXAQueryT0425(InstXAQuery):
    def __init__(self):
        self.make_blocks(0, 1)

    def get_order_info(self, account_number, is_executed="0", account_pwd=""):
        if is_executed == "체결":
            chegb = 1
        elif is_executed == "미체결":
            chegb = 2
        else:
            chegb = 0

        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "accno", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "passwd", 0, account_pwd)   # 비밀번호 없이 조회가능
        self.query.SetFieldData(self.in_block_list[0], "expcode", 0, "")    # 종목코드: 없으면 전체 조회
        self.query.SetFieldData(self.in_block_list[0], "chegb", 0, chegb)     # 체결구분: 전체(0), 체결(1), 미체결(2)
        self.query.SetFieldData(self.in_block_list[0], "medosu", 0, "0")    # 매매구분: 매도(1), 매수(2)
        self.query.SetFieldData(self.in_block_list[0], "sortgb", 0, "1")    # 정렬순서: 주문번호역순(1), 주문번호순(2)
        self.query.SetFieldData(self.in_block_list[0], "cts_ordno", 0, "")  # 연속 조회시 사용
        self.query.Request(False)

        ready_to_receive_query()

        order_info_list = []
        num_blocks = self.query.GetBlockCount(self.out_block_list[1])
        for i in range(num_blocks):
            ordno = self.query.GetFieldData(self.out_block_list[1], "ordno", i)
            expcode = self.query.GetFieldData(self.out_block_list[1], "expcode", i)
            medosu = self.query.GetFieldData(self.out_block_list[1], "medosu", i)   # 매매구분
            price = 0
            if chegb == 1:
                price = self.query.GetFieldData(self.out_block_list[1], "cheprice", i)     # 체결주문 조회시 체결가 표시
            else:
                price = self.query.GetFieldData(self.out_block_list[1], "price", i)  # 주문가격
            qty = self.query.GetFieldData(self.out_block_list[1], "qty", i)         # 주문수량
            cheqty = self.query.GetFieldData(self.out_block_list[1], "cheqty", i)   # 체결수량
            ordrem = self.query.GetFieldData(self.out_block_list[1], "ordrem", i)   # 미체결잔량
            ordtime = self.query.GetFieldData(self.out_block_list[1], "ordtime", i) # 주문시간
            ordtime = ordtime[0:2] + ":" + ordtime[2:4] + ":" + ordtime[4:6]
            # dictionary
            order_info = {"order_number": ordno, "stock_ticker": expcode, "order_type": medosu, "price": price,
                          "quantity": qty, "cheqty": cheqty, "ordrem": ordrem, "order_time": ordtime}
            order_info_list.append(order_info)

        return order_info_list


class InstXAQueryCSPAT00700(InstXAQuery):
    def __init__(self):
        self.make_blocks(1, 2)

    def change_order(self, account_number, account_pwd, stock_ticker, org_order_number, order_quantity, order_price, price_type="지정가"):
        ready_to_send_query()

        if price_type == "지정가":
            quote_type = "00"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, order_price)
        elif price_type == "시장가":
            quote_type = "03"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, 0)
        elif price_type == "종가":
            quote_type = "81"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, 0)
        elif price_type == "단일가":
            quote_type = "82"
            self.query.SetFieldData(self.in_block_list[0], "OrdPrc", 0, order_price)



        self.query.SetFieldData(self.in_block_list[0], "OrgOrdNo", 0, org_order_number)
        self.query.SetFieldData(self.in_block_list[0], "AcntNo", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "InptPwd", 0, account_pwd)
        self.query.SetFieldData(self.in_block_list[0], "IsuNo", 0, stock_ticker)
        self.query.SetFieldData(self.in_block_list[0], "OrdQty", 0, order_quantity)
        self.query.SetFieldData(self.in_block_list[0], "OrdprcPtnCode", 0, quote_type)  # 호가유형코드
        self.query.SetFieldData(self.in_block_list[0], "OrdCndiTpCode", 0, '0')  # 주문조건구분
        self.query.Request(False)

        ready_to_receive_query()

        order_result_list = []
        order_result_list.append(self.query.GetFieldData(self.out_block_list[1], "OrdNo", 0))
        order_time = self.query.GetFieldData(self.out_block_list[1], "OrdTime", 0)
        order_time = order_time[0:2] + ":" + order_time[2:4] + ":" + order_time[4:6]
        if order_result_list[0] == "0":
            return ["주문실패", False]
        order_result_list.append(order_time)
        return order_result_list


class InstXAQueryCSPAT00800(InstXAQuery):
    def __init__(self):
        self.make_blocks(1, 2)

    def cancel_order(self, account_number, account_pwd, stock_ticker, org_order_number, order_quantity):
        ready_to_send_query()
        
        self.query.SetFieldData(self.in_block_list[0], "AcntNo", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "InptPwd", 0, account_pwd)
        self.query.SetFieldData(self.in_block_list[0], "IsuNo", 0, stock_ticker)
        self.query.SetFieldData(self.in_block_list[0], "OrgOrdNo", 0, org_order_number)
        self.query.SetFieldData(self.in_block_list[0], "OrdQty", 0, order_quantity)
        self.query.Request(False)

        ready_to_receive_query()

        order_result_list = []
        order_result_list.append(self.query.GetFieldData(self.out_block_list[1], "OrdNo", 0))
        order_time = self.query.GetFieldData(self.out_block_list[1], "OrdTime", 0)
        order_time = order_time[0:2] + ":" + order_time[2:4] + ":" + order_time[4:6]
        if order_result_list[0] == "0":
            return ["주문실패", False]
        order_result_list.append(order_time)
        return order_result_list


# 주문가능금액 조회
class InstXAQueryCSPAQ12200(InstXAQuery):
    def __init__(self):
        self.make_blocks(1, 2)

    def get_amount_available_to_order(self, account_number, account_pwd=""):
        ready_to_send_query()

        self.query.SetFieldData(self.in_block_list[0], "RecCnt", 0, 1)
        self.query.SetFieldData(self.in_block_list[0], "AcntNo", 0, account_number)
        self.query.SetFieldData(self.in_block_list[0], "Pwd", 0, account_pwd)
        self.query.SetFieldData(self.in_block_list[0], "BalCreTp", 0, 0)
        self.query.SetFieldData(self.in_block_list[0], "MgmtBrnNo", 0, "")
        self.query.Request(False)

        ready_to_receive_query()

        return self.query.GetFieldData(self.out_block_list[1], "MnyOrdAbleAmt", 0)


class InstXAQueryT8430(InstXAQuery):
    def __init__(self):
        self.make_blocks()

    def get_all_stock_tickers(self):
        ready_to_send_query()
        self.query.SetFieldData(self.in_block_list[0], "gubun", 0, 0)
        self.query.Request(False)

        ready_to_receive_query()

        stock_ticker_list = []
        num_blocks = self.query.GetBlockCount(self.out_block_list[0])

        for i in range(num_blocks):
            shcode = self.query.GetFieldData(self.out_block_list[0], "shcode", i)
            hname = self.query.GetFieldData(self.out_block_list[0], "hname", i)
            stock_ticker = {"ticker": shcode, "name": hname}
            stock_ticker_list.append(stock_ticker)

        return pd.DataFrame(stock_ticker_list)


class InstXAQueryT8413(InstXAQuery):
    def __init__(self):
        self.make_blocks(0, 1)

    def get_chart_data(self, stock_ticker, quantity, end_date, start_date="", day_week_month=2):
        ready_to_send_query()
        self.query.SetFieldData(self.in_block_list[0], "shcode", 0, stock_ticker)
        self.query.SetFieldData(self.in_block_list[0], "gubun", 0, day_week_month)
        self.query.SetFieldData(self.in_block_list[0], "qrycnt", 0, quantity)
        self.query.SetFieldData(self.in_block_list[0], "sdate", 0, start_date)
        self.query.SetFieldData(self.in_block_list[0], "edate", 0, end_date)
        self.query.SetFieldData(self.in_block_list[0], "cts_date", 0, "")
        self.query.SetFieldData(self.in_block_list[0], "comp_yn", 0, "N")
        self.query.Request(False)

        ready_to_receive_query()

        chart_data_list = []
        for i in range(quantity):
            date = self.query.GetFieldData(self.out_block_list[1], "date", i)
            open_price = self.query.GetFieldData(self.out_block_list[1], "open", i)
            high_price = self.query.GetFieldData(self.out_block_list[1], "high", i)
            low_price = self.query.GetFieldData(self.out_block_list[1], "low", i)
            close_price = self.query.GetFieldData(self.out_block_list[1], "close", i)
            volume = self.query.GetFieldData(self.out_block_list[1], "jdiff_vol", i)
            chart_data = {"date": date, "open_price": open_price, "high_price": high_price, "low_price": low_price,
                          "close_price": close_price, "volume": volume}
            chart_data_list.append(chart_data)
        return  pd.DataFrame(chart_data_list)