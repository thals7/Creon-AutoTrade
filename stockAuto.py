import sys
import win32com.client  # COM 사용 위한 라이브러리
import ctypes  # 파이썬에서 DLL 로딩하여 DLL 제공 함수 호출 가능
import time
from datetime import datetime


cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')  # 크레온 상태 확인
cpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')  # 주문 오브젝트 사용 위한 초기화 작업
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')  # 주식 코드 조회
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')  # 주식 현재가 조회
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')  # 잔고 및 주문현황 평가 데이터
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')  # 매수 주문가능 금액/수량 조회
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  # 현금 주문 데이터 요청 및 수신


def printlog(message, *args):
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)


# 크레온 플러스 시스템 연결 상태 점검
def systemCheck():
    # 관리자 권한 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        printlog('Check Admin User : Failed')
        printlog('관리자 권한으로 실행하세요.')
        return False

    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        printlog('Check Server Connect : Failed')
        printlog('서버와의 연결이 끊겼습니다.')
        return False

    now = datetime.now()
    return print('크레온 플러스에 연결되었습니다. 현재시각은 {}월 {}일 {}시 {}분입니다.'.format(now.month, now.day, now.hour, now.minute))


# 인자로 받은 종목의 시장가, 매도호가, 매수호가 반환
def getStockCurr(code):
    cpStock.SetInputValue(0, code)
    cpStock.BlockRequest()

    item = {}
    item['현재가'] = cpStock.GetHeaderValue(11)  # 현재가
    item['매도호가'] = cpStock.GetHeaderValue(16)  # 매도호가
    item['매수호가'] = cpStock.GetHeaderValue(17)  # 매수호가
    return item


# 인자로 받은 종목의 전일 종가 반환
def getStockClosed(code):
    cpStock.SetInputValue(0, code)
    cpStock.BlockRequest()

    return cpStock.GetHeaderValue(10)


# 잔고 및 주문체결 평가 현황 조회
def getStockBalance():
    cpTrade.TradeInit()
    acc = cpTrade.AccountNumber[0]
    accFlag = cpTrade.GoodsList(acc, 1)  # 계좌 구분 (-1:전체 1:주식 2:선물/옵션)

    cpBalance.SetInputValue(0, acc)
    cpBalance.SetInputValue(1, accFlag[0])  # 주식 계좌중 첫번째 계좌
    cpBalance.BlockRequest()

    printlog('<계좌 종합 평가 현황>')
    printlog('계좌명 : {}'.format(cpBalance.GetHeaderValue(0)))
    # printlog('결제잔고수량 : {}'.format(cpBalance.GetHeaderValue(1)))
    printlog('종목수 : {}'.format(cpBalance.GetHeaderValue(7)))
    printlog('보유잔고수량 : {}'.format(cpBalance.GetHeaderValue(2)))
    # printlog('평가금액 : {0:,d} 원'.format((cpBalance.GetHeaderValue(3))))
    printlog('D+2 평가금액 : {0:,d} 원'.format((cpBalance.GetHeaderValue(9) + cpBalance.GetHeaderValue(11))))
    # printlog('잔고평가금액 : {0:,d} 원'.format((cpBalance.GetHeaderValue(11))))
    printlog('평가손익 : {0:,d} 원'.format((cpBalance.GetHeaderValue(4))))

    stocks = {}
    printlog('<개별 종목 평가 현황>')
    for i in range(cpBalance.GetHeaderValue(7)):
        code = cpBalance.GetDataValue(12, i)  # 종목코드
        name = cpBalance.GetDataValue(0, i)  # 종목명
        qty = cpBalance.GetDataValue(15, i)  # 수량
        val = cpBalance.GetDataValue(9, i)  # 평가금액
        pl = cpBalance.GetDataValue(10, i)  # 평가손익
        printlog('{} {}  수량 : {}  평가금액 : {:,d} 원  평가손익 : {:,d}원'.format(code, name, qty, val, pl))
        stocks[code] = {'name': name, 'qty': qty}
    return stocks


# 현금 주문 가능 금액 (D+2 예상 예수금)
def getCurrCash():
    cpTrade.TradeInit()
    acc = cpTrade.AccountNumber[0]
    accFlag = cpTrade.GoodsList(acc, 1)  # 계좌 구분 (-1:전체 1:주식 2:선물/옵션)

    cpBalance.SetInputValue(0, acc)
    cpBalance.SetInputValue(1, accFlag[0])  # 주식 계좌중 첫번째 계좌
    cpBalance.BlockRequest()
    return cpBalance.GetHeaderValue(9)


# 매수 주문
def buyOrder(code,qty):
    cpTrade.TradeInit()
    acc = cpTrade.AccountNumber[0]
    accFlag = cpTrade.GoodsList(acc, 1)  # 계좌 구분 (-1:전체 1:주식 2:선물/옵션)

    cpOrder.SetInputValue(0, '2')  # 2:매수
    cpOrder.SetInputValue(1, acc)
    cpOrder.SetInputValue(2, accFlag[0])  # 주식 상품중 첫번째
    cpOrder.SetInputValue(3, code)
    cpOrder.SetInputValue(4, qty)  # 매수할 수량
    cpOrder.SetInputValue(8, '03')  # 주문 호가 (3:시장가)
    cpOrder.BlockRequest()

    printlog('{} {}  수량 : {}'.format(code, cpCodeMgr.CodeToName(code), qty))
    printlog('매수 주문이 완료되었습니다.')


# 매도 주문
def sellOrder(code,qty):
    cpTrade.TradeInit()
    acc = cpTrade.AccountNumber[0]
    accFlag = cpTrade.GoodsList(acc, 1)  # 계좌 구분 (-1:전체 1:주식 2:선물/옵션)

    cpOrder.SetInputValue(0, '1')  # 1:매도
    cpOrder.SetInputValue(1, acc)
    cpOrder.SetInputValue(2, accFlag[0])
    cpOrder.SetInputValue(3, code)  # 종목코드
    cpOrder.SetInputValue(4, qty)  # 매도수량
    cpOrder.SetInputValue(8, '03')  # 주문 호가 (3:시장가)
    cpOrder.BlockRequest()

    printlog('{} {}  수량 : {}'.format(code, cpCodeMgr.CodeToName(code), qty))
    printlog('매도 주문이 완료되었습니다.')


if __name__ == '__main__':
    systemCheck()  # 크레온 접속 점검
    leverageList = ['A122630', 'A233740']
    inverseList = ['A114800','A251340']
    t_hour = [10, 11, 12, 13, 14, 15]
    timePercent = 0.15  # 시간당 매수 비율

    while True:
        t_now = datetime.now()
        t_sell = t_now.replace(hour=15, minute=20, second=0, microsecond=0)
        t_exit = t_now.replace(hour=15, minute=30, second=0, microsecond=0)
        weekday = datetime.today().weekday()
        if weekday == 5 or weekday == 6:  # 토요일이나 일요일이면 자동 종료
            printlog('오늘은 {}일 입니다. 프로그램이 종료됩니다.'.format('토요일' if weekday == 5 else '일요일'))
            sys.exit(0)

        if t_now.hour in t_hour and t_now.minute == 0:  # AM 10:00시부터 매시 정각마다 매수
            printlog('{}시 자동매매를 시작합니다.'.format(t_now.hour))
            stocks = getStockBalance()  # 현재 보유 종목 조회
            buyList = []  # 매수할 종목 리스트

            for stock in leverageList:
                if getStockCurr(stock)['현재가'] > getStockClosed(stock):  # 레버리지 ETF의 현재가보다 전일종가가 높으면 매수
                    buyList.append(stock)
            for stock in inverseList:
                if getStockCurr(stock)['현재가'] < getStockClosed(stock):  # 인버스 ETF의 현재가보다 전일종가가 낮으면 매수
                    buyList.append(stock)

            buyPercent = round(1/len(buyList),2)  # 종목당 매수 비율
            for buy in buyList:
                buyQty = int(getCurrCash()*timePercent*buyPercent/getStockCurr(buy)['현재가'])
                if buyQty < 1:
                    continue
                buyOrder(buy, buyQty)

            printlog('{}시 자동매수가 완료되었습니다.'.format(t_now.hour))
            time.sleep(5)

        getStockBalance()
        printlog('현재 D+2 예상 예수금은 {:,d} 원 입니다.'.format(getCurrCash()))

        if t_sell < t_now < t_exit:  # PM 03:20 ~ PM 03:30 : 일괄 매도
            printlog('{}월 {}일 일괄 매도를 시작합니다.'.format(t_sell.hour,t_sell.minute))
            for stock in getStockBalance():
                sellOrder(stock, stock['qty'])
            printlog('일괄 매도가 완료되었습니다.')
            time.sleep(20)

            getStockBalance()
            printlog('현재 D+2 예상 예수금은 {:,d} 원 입니다.'.format(getCurrCash()))
            time.sleep(600)

        if t_exit < t_now:  # PM 03:30 프로그램 종료
            printlog('장이 마감되었습니다. 프로그램을 종료합니다.')
            sys.exit(0)

        time.sleep(30)