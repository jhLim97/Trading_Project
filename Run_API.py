import sys
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *

TR_REQ_TIME_INTERVAL = 0.2

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.create_kiwoom_instance()
        self.set_signal_slots()

    def create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") #kiwoom's progID

    def set_signal_slots(self):
        self.OnEventConnect.connect(self.event_connect)
        self.OnReceiveTrData.connect(self.receive_tr_data)
        self.OnReceiveChejanData.connect(self.receive_chejan_data)

    def comm_connect(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

    def event_connect(self, err_code):
        if err_code == 0:
            print("connected")
        else:
            print("disconnected")

        self.login_event_loop.exit()

    def get_code_list_by_market(self, market):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market)
        code_list = code_list.split(';')
        return code_list[:-1]

    def get_master_code_name(self, code):
        code_name = self.dynamicCall("GetMasterCodeName(QString)", code)
        return code_name

    def get_connect_state(self):
        ret = self.dynamicCall("GetConnectState()")
        return ret

    def set_input_value(self,id,value):
        self.dynamicCall("SetInputValue(Qstring, Qstring)",id,value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall("CommRqData(QString, QString, int, QString", rqname, trcode, next, screen_no)
        self.tr_event_loop = QEventLoop()
        self.tr_event_loop.exec_()

    def get_comm_data(self, code, field_name, index, item_name):
        ret = self.dynamicCall("GetCommData(QString, QString, int, QString", code, field_name, index, item_name) #수정한 부분(더이상 지원안되는 메소드있어서:"commgetdata")
        return ret.strip()

    def get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def receive_tr_data(self, screen_no, rqname, trcode, record_name, next, unused1, unused2, unused3, unused4):
        if next == '2':
            self.remained_data = True

        else:
            self.remained_data = False

        if rqname == "opt10081_req":
            self.opt10081(rqname, trcode)

        elif rqname == "opw00001_req":
            self.opw00001(rqname, trcode)

        elif rqname == "opw00018_req":
            self.opw00018(rqname, trcode)

        try:
            self.tr_event_loop.exit()

        except AttributeError:
            pass


    def opt10081(self, rqname, trcode):
        data_cnt = self.get_repeat_cnt(trcode, rqname)

        for i in range(data_cnt):
            date = self.get_comm_data(trcode, rqname, i, "일자") #To_getcommdata
            open = self.get_comm_data(trcode, rqname, i, "시가")
            high = self.get_comm_data(trcode, rqname, i, "고가")
            low = self.get_comm_data(trcode, rqname, i, "저가")
            close = self.get_comm_data(trcode, rqname, i, "현재가")
            volume = self.get_comm_data(trcode, rqname, i, "거래량")

            self.ohlcv['date'].append(date)
            self.ohlcv['open'].append(int(open))
            self.ohlcv['high'].append(int(high))
            self.ohlcv['low'].append(int(low))
            self.ohlcv['close'].append(int(close))
            self.ohlcv['volume'].append(int(volume))

    def send_order(self, rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no):
        self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", [rqname, screen_no, acc_no, order_type, code, quantity, price, hoga, order_no])

    def get_chejan_data(self, fid):
        ret = self.dynamicCall("GetChejanData(int)", fid)
        return ret

    def receive_chejan_data(self, gubun, item_cnt, fid_list):
        print(gubun)
        print(self.get_chejan_data(9203))
        print(self.get_chejan_data(302))
        print(self.get_chejan_data(900))
        print(self.get_chejan_data(901))

    def get_login_info(self, tag):
        ret = self.dynamicCall("GetLoginInfo(QString)", tag)
        return ret

    @staticmethod #checking
    def change_format(data):
        strip_data = data.lstrip('-0')
        if strip_data == '':
            strip_data = '0'

        try:
            format_data = format(int(strip_data), ',d')
        except:
            format_data = format(float(strip_data))

        if data.startswith('-'):
            format_data = '-' + format_data

        return format_data

    @staticmethod
    def change_format2(data):
        strip_data = data.lstrip('-0')

        if strip_data == '':
            strip_data = '0'

        if strip_data.startswith('.'):
            strip_data = '0' + strip_data

        if data.startswith('-'):
            strip_data = '-' + strip_data

        return strip_data

    def opw00001(self, rqname, trcode):
        d2_deposit = self.get_comm_data(trcode, rqname, 0, "d+2추정예수금") #getcommdata로 변경.
        self.d2_deposit = Kiwoom.change_format(d2_deposit)

    def opw00018(self, rqname, trcode):
        total_purchase_price = self.get_comm_data(trcode, rqname, 0, "총매입금액") #To_getcommdata
        total_eval_price = self.get_comm_data(trcode, rqname, 0, "총평가금액")
        total_eval_profit_loss_price = self.get_comm_data(trcode, rqname, 0, "총평가손익금액")
        total_earning_rate = self.get_comm_data(trcode, rqname, 0, "총수익률(%)")
        estimated_deposit = self.get_comm_data(trcode, rqname, 0, "추정예탁자산")

        # 실 서버와 모의투자 서버로 접속 시 제공되는 데이터의 형식다름. 모의는 소수점포함. 이를 위한 조정메서드.
        total_earning_rate = Kiwoom.change_format(total_earning_rate)
        if self.get_server_gubun():
            total_earning_rate = float(total_earning_rate) / 100
            total_earning_rate = str(total_earning_rate)

        total_purchase_price = Kiwoom.change_format(total_purchase_price)
        total_eval_price = Kiwoom.change_format(total_eval_price)
        total_eval_profit_loss_price = Kiwoom.change_format(total_eval_profit_loss_price)
        estimated_deposit = Kiwoom.change_format(estimated_deposit)
        self.opw00018_output['single'].append([total_purchase_price, total_eval_price, total_eval_profit_loss_price, total_earning_rate, estimated_deposit])

        # multi data ,opw00018의 경우 한 번에 최대 20개의 보유 종목에 대한 데이터 얻을 수 있음.
        rows = self.get_repeat_cnt(trcode, rqname) #이 부분 모의매매 후 테스팅 해볼 것.
        for i in range(rows):
            issue_name = self.get_comm_data(trcode, rqname, i, "종목명") #To_getcommdata
            quantity = self.get_comm_data(trcode, rqname, i, "보유수량")
            purchase_price = self.get_comm_data(trcode, rqname, i, "매입가")
            current_price = self.get_comm_data(trcode, rqname, i, "현재가")
            eval_profit_loss_price = self.get_comm_data(trcode, rqname, i, "평가손익")
            earning_rate = self.get_comm_data(trcode, rqname, i, "수익률(%)")

            quantity = Kiwoom.change_format(quantity)
            purchase_price = Kiwoom.change_format(purchase_price)
            current_price = Kiwoom.change_format(current_price)
            eval_profit_loss_price = Kiwoom.change_format(eval_profit_loss_price)
            earning_rate = Kiwoom.change_format2(earning_rate)

            self.opw00018_output['multi'].append([issue_name, quantity, purchase_price, current_price, eval_profit_loss_price, earning_rate])

    def reset_opw00018_output(self):
        self.opw00018_output = {'single': [], 'multi': []}

    def get_server_gubun(self): #3일차 2)마지막 실 서버와 모의투자 서버로 접속 시 제공되는 데이터의 형식다름. 모의는 소수점포함. 이를 위한 조정메서드.
        ret = self.dynamicCall("KOA_Functions(QString, QString)", "GetServerGubun", "")
        return ret