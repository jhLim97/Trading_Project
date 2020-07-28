from PyQt5 import uic
from Run_API import *
from Soaring_stock import *
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import pandas_datareader.data as web
from datetime import datetime
import time
import pandas as pd
from pandas import Series, DataFrame

UI_FORM = uic.loadUiType("Trader.ui")[0]

class SubWindow(QDialog):
    def __init__(self,parent):
        super(SubWindow, self).__init__(parent)
        sub_ui = 'Chart.ui'
        uic.loadUi(sub_ui, self)
        self.show()
        self.making_window()

    def making_window(self):
        self.pushButton.clicked.connect(self.drawing_chart)

        self.fig = plt.Figure()
        self.canvas = FigureCanvas(self.fig)
        self.verticalLayout.addWidget(self.canvas)

        self.verticalLayout_2.addStretch(1)

    def drawing_chart(self):
        code = self.lineEdit.text()

        year = datetime.today().year
        month = datetime.today().month
        day = datetime.today().day

        if month < 5:
            start_year = year - 1
            start_month = month + 8

        else:
            start_year = year
            start_month = month - 4

        _date = str(start_year) + '-' + str(start_month) + '-' + '01' #지금으로부터 4개월전 해당 달의 1일부터.
        start = datetime.strptime(_date,"%Y-%m-%d").date()
        date_ = str(year) + '-' + str(month) + '-' + str(day) #현재까지.
        end = datetime.strptime(date_,"%Y-%m-%d").date()

        df = web.DataReader(code + ".KS", "yahoo", start, end) #특정 구간 비교 원할경우 이부분에 start,end명시.
        df['MA20'] = df['Adj Close'].rolling(window=20).mean()
        df['MA60'] = df['Adj Close'].rolling(window=60).mean()
        df['MA120'] = df['Adj Close'].rolling(window=120).mean()

        ax = self.fig.add_subplot(111)
        ax.plot(df.index, df['Adj Close'], label='Adj Close')
        ax.plot(df.index, df['MA20'], label='MA20')
        ax.plot(df.index, df['MA60'], label='MA60')
        ax.plot(df.index, df['MA120'], label='MA120')
        ax.legend(loc='upper left')
        ax.grid()

        self.canvas.draw()


class MyWindow(QMainWindow, UI_FORM):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.trade_stocks_done = False

        self.kiwoom = Kiwoom()
        self.kiwoom.comm_connect()

        self._times = Times()

        self.pushButton_3.clicked.connect(self.load_subwindow)
        self.pushButton_4.clicked.connect(self.Initialize)
        self.pushButton_5.clicked.connect(self.several_times_control)

        #Timer1
        self.timer = QTimer(self)
        self.timer.start(1000) #1초에 한 번 갱신.
        self.timer.timeout.connect(self.timeout)

        self.lineEdit.textChanged.connect(self.code_changed)
        self.pushButton_2.clicked.connect(self.check_balance)

        # Timer2
        self.timer2 = QTimer(self)
        self.timer2.start(1000 * 3) #3초에 한 번 갱신.
        self.timer2.timeout.connect(self.timeout2)

        accounts_num = int(self.kiwoom.get_login_info("ACCOUNT_CNT"))
        accounts = self.kiwoom.get_login_info("ACCNO")
        accounts_list = accounts.split(';')[0:accounts_num]
        self.comboBox_2.addItems(accounts_list) #comboBox 객체이름 내가 섞이게함(일반적인 상황에서는 "comboBox").

        self.pushButton.clicked.connect(self.send_order)

        self.load_buy_sell_list() #파일로부터 정보읽어들이기.

    def several_times_control(self):
        times = self.spinBox_3.value()
        self._times.decision_times(times)


    def load_subwindow(self):
        SubWindow(self)

    def timeout(self):
        market_start_time = QTime(9, 0, 0)
        current_time = QTime.currentTime()

        if current_time > market_start_time and self.trade_stocks_done is False:
            self.trade_stocks()
            self.trade_stocks_done = True

        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        state = self.kiwoom.GetConnectState()
        if state == 1:
            state_msg = "서버 연결 중"
        else:
            state_msg = "서버 미 연결 중"

        self.statusbar.showMessage(state_msg + " | " + time_msg)

    def code_changed(self):
        code = self.lineEdit.text()
        name = self.kiwoom.get_master_code_name(code)
        self.lineEdit_2.setText(name)

    def send_order(self):
        order_type_lookup = {'신규매수': 1, '신규매도': 2, '매수취소': 3, '매도취소': 4} #키움API보고 추가하기
        hoga_lookup = {'지정가': "00", '시장가': "03"} #키움API보고 추가하기

        account = self.comboBox_2.currentText() #COMBOBOX순서 내가 섞어놈.
        order_type = self.comboBox.currentText()
        code = self.lineEdit.text()
        hoga = self.comboBox_3.currentText()
        num = self.spinBox.value()
        price = self.spinBox_2.value()

        self.kiwoom.send_order("send_order_req", "0101", account, order_type_lookup[order_type], code, num, price, hoga_lookup[hoga], "")

    def check_balance(self):
        self.kiwoom.reset_opw00018_output()
        account_number = self.kiwoom.get_login_info("ACCNO")
        account_number = account_number.split(';')[0]

        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 0, "2000")

        while self.kiwoom.remained_data: #한 번에 최대 20개 가져오므로 남은 데이터 처리.
            time.sleep(0.2)
            self.kiwoom.set_input_value("계좌번호", account_number)
            self.kiwoom.comm_rq_data("opw00018_req", "opw00018", 2, "2000")

        # opw00001
        self.kiwoom.set_input_value("계좌번호", account_number)
        self.kiwoom.comm_rq_data("opw00001_req", "opw00001", 0, "2000")

        # balance
        item = QTableWidgetItem(self.kiwoom.d2_deposit) # QTableWidget에 출력하기 위해 먼저 self.kiwoom.d2_deposit에 저장된 예수금 데이터를 QTableWidgetItem 객체로 변환.
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.tableWidget.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw00018_output['single'][0][i-1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget.setItem(0, i, item)
        self.tableWidget.resizeRowsToContents() #아이템의 크기에 맞춰 행의 높이 조절.

        # Item list
        item_count = len(self.kiwoom.opw00018_output['multi']) #앞에 tablewidget도 이런식으로 소스파일에서 코딩해보기.
        self.tableWidget_2.setRowCount(item_count)

        for i in range(item_count):
            row = self.kiwoom.opw00018_output['multi'][i]
            for j in range(len(row)):
                item = QTableWidgetItem(row[j])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.tableWidget_2.setItem(i, j, item)
        self.tableWidget_2.resizeRowsToContents()

    def timeout2(self):
        if self.checkBox.isChecked():
            self.check_balance()

    def load_buy_sell_list(self):
        f = open("buy_list.txt", 'rt')
        buy_list = f.readlines()
        f.close()

        f = open("sell_list.txt", 'rt', encoding='UTF8') #UnicodeError 발생 시 "encoding='UTF8'" 이 문장 지우고 실행해보기.
        sell_list = f.readlines()
        f.close()

        row_count = len(buy_list) + len(sell_list)
        self.tableWidget_3.setRowCount(row_count)

        # buy list
        for i in range(len(buy_list)):
            row_data = buy_list[i]
            split_row_data = row_data.split(';')
            split_row_data[1] = self.kiwoom.get_master_code_name(split_row_data[1].rsplit())

            for j in range(len(split_row_data)):
                item = QTableWidgetItem(split_row_data[j].rstrip())
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignCenter)
                self.tableWidget_3.setItem(i, j, item)

        # sell list
        for i in range(len(sell_list)):
            row_data = sell_list[i]
            split_row_data = row_data.split(';')
            split_row_data[1] = self.kiwoom.get_master_code_name(split_row_data[1].rstrip())

            for j in range(len(split_row_data)):
                item = QTableWidgetItem(split_row_data[j].rstrip())
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignCenter)
                self.tableWidget_3.setItem(len(buy_list) + i, j, item)

        self.tableWidget_3.resizeRowsToContents()

    def trade_stocks(self):
        hoga_lookup = {'지정가': "00", '시장가': "03"}

        f = open("buy_list.txt", 'rt')
        buy_list = f.readlines()
        f.close()

        f = open("sell_list.txt", 'rt', encoding='UTF8') #UnicodeError 발생 시 "encoding='UTF8'" 이 문장 지우고 실행해보기.
        sell_list = f.readlines()
        f.close()

        account = self.comboBox_2.currentText() #2인 이유는 내가 순서 바꿔놓음.

        # buy list
        for row_data in buy_list:
            split_row_data = row_data.split(';')
            hoga = split_row_data[2]
            code = split_row_data[1]
            num = split_row_data[3]
            price = split_row_data[4]

            if split_row_data[-1].rstrip() == '매수 전': #“RQ_1”으로 보내도됨?
                self.kiwoom.send_order("send_order_req", "0101", account, 1, code, num, price, hoga_lookup[hoga], "")

        # sell list
        for row_data in sell_list:
            split_row_data = row_data.split(';')
            hoga = split_row_data[2]
            code = split_row_data[1]
            num = split_row_data[3]
            price = split_row_data[4]

            if split_row_data[-1].rstrip() == '매도 전': #“RQ_1”으로 보내도됨?
                self.kiwoom.send_order("send_order_req", "0101", account, 2, code, num, price, hoga_lookup[hoga], "")

        # buy list
        for i, row_data in enumerate(buy_list):
            buy_list[i] = buy_list[i].replace("매수 전", "주문완료") #row_data는 사용안함?

        # file update
        f = open("buy_list.txt", 'wt')
        for row_data in buy_list:
            f.write(row_data)
        f.close()

        # sell list
        for i, row_data in enumerate(sell_list):
            sell_list[i] = sell_list[i].replace("매도 전", "주문완료")

        # file update
        f = open("sell_list.txt", 'wt', encoding='UTF8') #UnicodeError 발생 시 "encoding='UTF8'" 이 문장 지우고 실행해보기.
        for row_data in sell_list:
            f.write(row_data)
        f.close()

    def Initialize(self):
        self.tableWidget.clearContents()
        self.tableWidget_2.clearContents()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()