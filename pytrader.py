from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import uic
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from operator import itemgetter
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import pandas as pd
import time
import sys
from ebest import *


form_class = uic.loadUiType("PyTrader.ui")[0]


def set_font(font_dir):
    font_family = fm.FontProperties(fname=font_dir).get_name()
    plt.rcParams["font.family"] = font_family


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.instXASession = InstXASession()
        self.display_account_code()

        # 쿼리 객체 만들기
        self.instXAQueryT1102 = InstXAQueryT1102()
        self.instXAQueryT0424 = InstXAQueryT0424()
        self.instXAQueryT0425 = InstXAQueryT0425()
        self.instXAQueryCSPAT00600 = InstXAQueryCSPAT00600()    # 정상주문
        self.instXAQueryCSPAT00700 = InstXAQueryCSPAT00700()    # 정정주문
        self.instXAQueryCSPAT00800 = InstXAQueryCSPAT00800()    # 취소주문

        self.orderNumberLine.setDisabled(True)

        # 클릭 이벤트

        self.tickerCodeLine.textChanged.connect(self.display_ticker_name)
        self.orderButton.clicked.connect(self.display_order_result)
        self.balanceInquiryButton.clicked.connect(self.display_account_balance_info)
        self.orderInquiryButton.clicked.connect(self.display_account_order_info)
        self.orderTypeComboBox.currentIndexChanged.connect(self.change_ui_visibility_based_on_order_type)
        self.unconOrderTableWidget.itemSelectionChanged.connect(self.one_click_set_correct_order)
        self.stocksTableWidget.itemSelectionChanged.connect(self.one_click_set_normal_order)
        self.realtimeCheckBox.stateChanged.connect(self.realtime_inquiry_checked)
        self.portfolioInquiryButton.clicked.connect(self.display_portfolio)

        self.timer1 = QTimer(self)
        self.timer1.timeout.connect(self.display_account_balance_info)
        
        # 멤버 변수
        self.account_info_list = []
        self.holding_stocks_list = []

        # 포트폴리오 파이그래프
        self.fig = plt.Figure()
        self.canvas = FigureCanvas(self.fig)
        self.canvasLayout.addWidget(self.canvas)


    def display_account_code(self):
        self.accountNumberComboBox.addItems(self.instXASession.get_account_code())

    def change_ui_visibility_based_on_order_type(self):
        current_index = self.orderTypeComboBox.currentIndex()
        if current_index == 0 or current_index == 1:
            self.orderNumberLine.setDisabled(True)
            self.tickerPriceSpinBox.setDisabled(False)
            self.priceTypeComboBox.setDisabled(False)
        elif current_index == 2:
            self.orderNumberLine.setDisabled(False)
            self.tickerPriceSpinBox.setDisabled(False)
            self.priceTypeComboBox.setDisabled(False)
        elif current_index == 3:
            self.orderNumberLine.setDisabled(False)
            self.tickerPriceSpinBox.setDisabled(True)
            self.priceTypeComboBox.setDisabled(True)

    def display_ticker_name(self):
        ticker_code = self.tickerCodeLine.text()
        self.tickerNameLine.setText(self.instXAQueryT1102.get_name_of_ticker(ticker_code))
        self.tickerPriceSpinBox.setValue(self.instXAQueryT1102.get_price_of_ticker(ticker_code))
        if self.tickerPriceSpinBox.value() != 0:
            self.tickerQuantitySpinBox.setValue(1)

    def display_order_result(self):
        order_type = str(self.orderTypeComboBox.currentText())
        account_pwd = self.get_account_pwd_from_user()
        status_bar_message = {}
        if order_type == "신규매수" or order_type == "신규매도":
            order_result_list = self.instXAQueryCSPAT00600.send_order(str(self.accountNumberComboBox.currentText()), str(account_pwd),
                                                                      str(self.tickerCodeLine.text()), int(self.tickerQuantitySpinBox.value()),
                                                                      int(self.tickerPriceSpinBox.value()), order_type,
                                                                      str(self.priceTypeComboBox.currentText()))
            status_bar_message = {"주문종류": order_type, "주문번호": order_result_list[0], "주문시각": order_result_list[1]}
            self.statusBar().showMessage(str(status_bar_message))

        elif order_type == "정정주문":
            order_result_list = self.instXAQueryCSPAT00700.change_order(str(self.accountNumberComboBox.currentText()), str(account_pwd),
                                                                        str(self.tickerCodeLine.text()), str(self.orderNumberLine.text()),
                                                                        int(self.tickerQuantitySpinBox.value()), int(self.tickerPriceSpinBox.value()),
                                                                        str(self.priceTypeComboBox.currentText()))
            status_bar_message = {"주문종류": order_type, "주문번호": order_result_list[0], "주문시각": order_result_list[1]}
            self.statusBar().showMessage(str(status_bar_message))

        elif order_type == "취소주문":
            order_result_list = self.instXAQueryCSPAT00800.cancel_order(str(self.accountNumberComboBox.currentText()), str(account_pwd),
                                                                   str(self.tickerCodeLine.text()), str(self.orderNumberLine.text()),
                                                                   int(self.tickerQuantitySpinBox.value()))
            status_bar_message = {"주문종류": order_type, "주문번호": order_result_list[0], "주문시각": order_result_list[1]}
            self.statusBar().showMessage(str(status_bar_message))
        if status_bar_message["주문번호"] == "주문실패":
            self.statusBar().setStyleSheet("background-color: #f06060")
        else:
            self.statusBar().setStyleSheet("background-color: #b0ffb7")
        time.sleep(0.2)
        self.display_account_balance_info()
        self.display_account_order_info()
        time.sleep(0.5)
        self.statusBar().setStyleSheet("background-color: #f4fff3")

    def display_account_balance_info(self):
        #account_pwd = self.get_account_pwd_from_user()
        self.account_info_list = self.instXAQueryT0424.get_account_balance_info(str(self.accountNumberComboBox.currentText()))
        len_of_list = len(self.account_info_list)
        self.balanceTableWidget.setRowCount(len_of_list)  
        
        for i in range(len_of_list):
            j = 0
            for key in self.account_info_list[i]:
                info = QTableWidgetItem(self.account_info_list[i][key])
                info.setTextAlignment(Qt.AlignCenter | Qt.AlignRight)
                self.balanceTableWidget.setItem(i, j, info)
                j += 1
                
        # 계좌가 1개라고 가정 상태임
    
        self.holding_stocks_list = self.instXAQueryT0424.get_holding_stocks_info(str(self.accountNumberComboBox.currentText()))
        len_of_list = len(self.holding_stocks_list)
        self.stocksTableWidget.setRowCount(len_of_list)

        for i in range(len_of_list):
            j = 0
            for key in self.holding_stocks_list[i]:
                info = QTableWidgetItem(self.holding_stocks_list[i][key])
                if key == "profit_rate":
                    if float(info.text()) > 0:
                        info.setForeground(QBrush(QColor(255, 0, 0)))
                    elif float(info.text()) < 0:
                        info.setForeground(QBrush(QColor(0, 0, 255)))
                info.setTextAlignment(Qt.AlignCenter | Qt.AlignRight)
                self.stocksTableWidget.setItem(i, j, info)
                j += 1

    def display_account_order_info(self):
        # 마지막 파라미터는 체결 미체결 구분
        order_info_list = self.instXAQueryT0425.get_order_info(str(self.accountNumberComboBox.currentText()), "체결")
        len_of_list = len(order_info_list)
        self.conOrderTableWidget.setRowCount(len_of_list)

        for i in range(len_of_list):
            j = 0
            for key in order_info_list[i].keys():
                info = QTableWidgetItem(order_info_list[i][key])
                if key == "order_type":
                    if info.text()[0:2] == "매수":
                        info.setForeground(QBrush(QColor(255, 0, 0)))
                    elif info.text()[0:2] == "매도":
                        info.setForeground(QBrush(QColor(0, 0, 255)))

                info.setTextAlignment(Qt.AlignCenter | Qt.AlignRight)
                self.conOrderTableWidget.setItem(i, j, info)
                j += 1
                
        order_info_list = self.instXAQueryT0425.get_order_info(str(self.accountNumberComboBox.currentText()), "미체결")
        len_of_list = len(order_info_list)
        self.unconOrderTableWidget.setRowCount(len_of_list)

        for i in range(len_of_list):
            j = 0
            for key in order_info_list[i].keys():
                info = QTableWidgetItem(order_info_list[i][key])
                if key == "order_type":
                    if info.text()[0:2] == "매수":
                        info.setForeground(QBrush(QColor(255, 0, 0)))
                    elif info.text()[0:2] == "매도":
                        info.setForeground(QBrush(QColor(0, 0, 255)))
                info.setTextAlignment(Qt.AlignCenter | Qt.AlignRight)
                self.unconOrderTableWidget.setItem(i, j, info)
                j += 1

    def one_click_set_correct_order(self):
        try:
            current_row = self.unconOrderTableWidget.currentRow()
            self.orderTypeComboBox.setCurrentIndex(2)
            self.orderNumberLine.setText(self.unconOrderTableWidget.item(current_row, 0).text())
            self.tickerCodeLine.setText(self.unconOrderTableWidget.item(current_row, 1).text())
            self.tickerQuantitySpinBox.setValue(int(self.unconOrderTableWidget.item(current_row, 6).text()))

        except AttributeError:
            return

    def one_click_set_normal_order(self):
        try:
            current_row = self.stocksTableWidget.currentRow()
            self.orderTypeComboBox.setCurrentIndex(0)
            self.tickerCodeLine.setText(self.stocksTableWidget.item(current_row, 1).text())

        except AttributeError:
            return

    def get_account_pwd_from_user(self):
        account_pwd, ok = QInputDialog.getText(self, "주의", "계좌비밀번호를 입력하세요.", QLineEdit.Password)
        if ok is False:
            return "----"
        return account_pwd

    def realtime_inquiry_checked(self):
        if self.realtimeCheckBox.checkState() == 2:
            self.display_account_balance_info()
            self.timer1.start(1000 * 10)
        elif self.realtimeCheckBox.checkState() == 0:
            self.timer1.stop()

    def display_portfolio(self):
        total_estimated_amount = 0  # 모든 계좌의 추정순자산 합
        stocks_to_be_drawn_list = []         # 그래프로 나타낼 주식들
        if len(self.account_info_list) == 0:    # 계좌 데이터가 없을 경우 불러오기
            self.display_account_balance_info()

        for i in range(len(self.account_info_list)):
            total_estimated_amount += int(self.account_info_list[i]["estimated_amount"])

        self.totalAmountLabel.setText("[ 총자산 : " + str(total_estimated_amount) + "원 ]")

        for i in range(len(self.holding_stocks_list)):
            quantity = int(self.holding_stocks_list[i]["quantity"])
            if quantity != 0:
                estimated_amount = quantity * int(self.holding_stocks_list[i]["cur_price"])
                total_estimated_amount -= estimated_amount
                stocks_to_be_drawn = {"ticker_name": self.holding_stocks_list[i]["ticker_name"], "estimated_amount": estimated_amount}
                stocks_to_be_drawn_list.append(stocks_to_be_drawn)
        stocks_to_be_drawn_list.append({"ticker_name": "현금", "estimated_amount": total_estimated_amount})  # 주식을 제외한 보유 현금

        stocks_to_be_drawn_list = sorted(stocks_to_be_drawn_list, key=itemgetter("estimated_amount", "ticker_name"), reverse=True)

        ratio = []
        labels = []
        for i in range(len(stocks_to_be_drawn_list)):
            ratio.append(stocks_to_be_drawn_list[i]["estimated_amount"])
            labels.append(stocks_to_be_drawn_list[i]["ticker_name"])

        colors = ['#ff9999', '#ffc000', '#dbff8b', '#8fd9b6', '#9ed0ff', '#d395d0', '#7f7bff']
        self.fig.clear()
        self.fig.add_subplot().pie(ratio, labels=labels, autopct='%.1f%%', startangle=260, counterclock=False,
                                   colors=colors)
        self.canvas.draw()

        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()