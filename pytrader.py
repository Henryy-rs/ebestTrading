import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5 import uic
from ebest import *
import time

form_class = uic.loadUiType("PyTrader.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.instXASession = InstXASession()
        self.display_account_code()

        # 쿼리 객체 만들기
        self.instXAQueryT1102 = InstXAQueryT1102()
        self.instXAQueryT0424 = InstXAQueryT0424()
        self.instXAQueryCSPAT00600 = InstXAQueryCSPAT00600()

        #self.instXAQueryCSPAQ12200 = InstXAQueryCSPAQ12200()
        #self.instXAQueryCSPAQ12300 = InstXAQueryCSPAQ12300()

        self.tickerCodeLine.textChanged.connect(self.when_ticker_code_is_changed)
        self.orderButton.clicked.connect(self.order_button_clicked)
        self.inquiryButton.clicked.connect(self.check_account_button_clicked)

    def display_account_code(self):
        self.accountNumberComboBox.addItems(self.instXASession.get_account_code())

    def when_ticker_code_is_changed(self):
        ticker_code = self.tickerCodeLine.text()
        self.tickerNameLine.setText(self.instXAQueryT1102.get_name_of_ticker(ticker_code))
        self.tickerPriceSpinBox.setValue(self.instXAQueryT1102.get_price_of_ticker(ticker_code))
        if self.tickerPriceSpinBox.value() != 0:
            self.tickerAmountSpinBox.setValue(1)

    def order_button_clicked(self):
        account_pwd = self.get_account_pwd_from_user()
        self.instXAQueryCSPAT00600.send_order(str(self.accountNumberComboBox.currentText()), str(account_pwd),
                                              str(self.tickerCodeLine.text()), int(self.tickerAmountSpinBox.value()),
                                              int(self.tickerPriceSpinBox.value()),
                                              str(self.orderTypeComboBox.currentText()),
                                              str(self.priceTypeComboBox.currentText()))

    def check_account_button_clicked(self):
        account_pwd = self.get_account_pwd_from_user()
        account_info_list = self.instXAQueryT0424.get_account_balance(str(self.accountNumberComboBox.currentText()),
                                                                      str(account_pwd))
        len_of_list = len(account_info_list)
        self.balanceTableWidget.setRowCount(1)  # 계좌가 1개라고 가정함

        for i in range(len_of_list):
            info = QTableWidgetItem(account_info_list[i])
            info.setTextAlignment(Qt.AlignCenter | Qt.AlignRight)
            self.balanceTableWidget.setItem(0, i, info)

        holding_stocks_list = self.instXAQueryT0424.get_holding_stocks(str(self.accountNumberComboBox.currentText()),
                                                                      str(account_pwd))
        len_of_list = len(holding_stocks_list)
        self.stocksTableWidget.setRowCount(len_of_list)

        for i in range(len_of_list):
            for j in range(len(holding_stocks_list[i])):
                info = QTableWidgetItem(holding_stocks_list[i][j])
                info.setTextAlignment(Qt.AlignCenter | Qt.AlignRight)
                self.stocksTableWidget.setItem(i, j, info)

    def get_account_pwd_from_user(self):
        account_pwd, ok = QInputDialog.getText(self, "주의", "계좌비밀번호를 입력하세요.", QLineEdit.Password)
        if ok is False:
            return "----"
        return account_pwd


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()