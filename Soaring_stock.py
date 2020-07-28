import sys
from PyQt5.QtWidgets import *
import Run_API
import time
from pandas import DataFrame
import datetime

MARKET_KOSPI = 0
MARKET_KOSDAQ = 10
TIMES = 1 #초기 급등주 기준 : 오늘 이전 20일동안의 평균의 1배수.

class Times:
    def decision_times(self, times):
        global TIMES
        TIMES = times

class Soaring:
    def __init__(self):
        self.kiwoom = Run_API.Kiwoom()
        self.kiwoom.comm_connect()
        self.get_code_list()

    def get_code_list(self):
        self.kospi_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSPI)
        self.kosdaq_codes = self.kiwoom.get_code_list_by_market(MARKET_KOSDAQ)

    def Run(self):
        buy_list = []
        num = len(self.kosdaq_codes)

        for i, code in enumerate(self.kosdaq_codes):
            print(i,'/',num)
            if self.check_speedy_rising_volume(code):
                buy_list.append(code)

        self.update_buy_list((buy_list))

    def get_ohlcv(self, code, start):
        self.kiwoom.ohlcv = {'date': [], 'open': [], 'high': [], 'low': [], 'close': [], 'volume': []}
        self.kiwoom.set_input_value("종목코드", code)
        self.kiwoom.set_input_value("기준일자", start)
        self.kiwoom.set_input_value("수정주가구분", 1)
        self.kiwoom.comm_rq_data("opt10081_req", "opt10081", 0, "0101")

        time.sleep(0.2)

        df = DataFrame(self.kiwoom.ohlcv, columns=['open', 'high', 'low', 'close', 'volume'], index=self.kiwoom.ohlcv['date'])
        return df

    def check_speedy_rising_volume(self, code):
        today = datetime.datetime.today().strftime("%Y%m%d")
        df = self.get_ohlcv(code, today)
        volumes = df['volume'] #거래량 칼럼만 바인딩.

        if len(volumes) < 21:
            return False

        sum_vol20 = 0
        today_vol = 0

        for i, vol in enumerate(volumes):
            if i == 0:
                today_vol = vol
            elif 1 <= i <= 20:
                sum_vol20 += vol
            else:
                break

        avg_vol20 = sum_vol20 / 20
        if today_vol >= avg_vol20 * int(TIMES): #급등주 기준 수정시 이 부분 변경.
            return True

    def update_buy_list(self, buy_list):
        f = open("buy_list.txt", "wt")
        for code in buy_list:
            f.writelines("매수;%s;시장가;10;0;매수 전\n" % (code))
        f.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    soaring = Soaring()
    soaring.Run()