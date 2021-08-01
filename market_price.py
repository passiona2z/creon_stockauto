import win32com.client
import requests


# 크레온 시스템 체크
def check_creon_system():

    objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
    bConnect = objCpCybos.IsConnect
    if (bConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        exit()


# 크레온 시스템을 통해 정보 불러오고 반환
def get_current_info(code):
    
    cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()

    item = {}
    item['name']= cpStock.GetHeaderValue(1)   # 종목명
    item['c_price'] = cpStock.GetHeaderValue(11)   # 종가
    item['diff'] =  cpStock.GetHeaderValue(12)        # 대비
    item['vol'] =  cpStock.GetHeaderValue(18)        # 거래량  
    return item['name'], item['c_price'], item['diff'], item['vol']

# 슬랙 메세지 보내기
def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text})
    print(response)


# 주식 리스트 # 삼성전자 본주 및 우선주 # LG생활건강 본주 및 우선주
stock_list = ['A005930', 'A005935', 'A051900', 'A051905']


# 함수 실행 (주식코드별 반복)
for i in stock_list :
    check_creon_system()
    name, c_price, diff, vol = get_current_info(i) 
 
    myToken = "xoxb-2331042902405-2334084254274-w79unP0ckGKK5yK73heOu7kQ"
    post_message(myToken,"#real-project","{} > 종가 : {:,} / 대비 : {:,} / 거래량 : {:,}" .format(name, c_price, diff, vol))

    print("{} > 종가 : {:,} / 대비 : {:,} / 거래량 : {:,}" .format(name, c_price, diff, vol)) 
    
                               
