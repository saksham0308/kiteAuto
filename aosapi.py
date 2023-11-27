# # package import statement
import SmartApi
from SmartApi import SmartConnect  # or from SmartApi.smartConnect import SmartConnect
import pyotp
import datetime
import time
print("Got smartConnect APi key at")
m=datetime.datetime.now()
print(m.second)

api_key = 'qhzBkFCL'
clientId = 'S52797606'
pwd = '9717'
smartApi = SmartConnect(api_key)

print("Got smartConnect APi key at")
m=datetime.datetime.now()
print(m.second)

token = "VQNBZD2P4ZQLBAE22WGMBRHXCI"
totp = pyotp.TOTP(token).now()

print("Got TOTP Token at")
m=datetime.datetime.now()
print(m.second)

correlation_id = "abc123"

# login api call

data = smartApi.generateSession(clientId, pwd, totp)
# print(data)
authToken = data['data']['jwtToken']

print("Got authToken at")
m=datetime.datetime.now()
print(m.second)


refreshToken = data['data']['refreshToken']

print("Got refreshToken at")
m=datetime.datetime.now()
print(m.second)


# fetch the feedtoken
feedToken = smartApi.getfeedToken()

print("Got feedToken at")
m=datetime.datetime.now()
print(m.second)

# fetch User Profile
res = smartApi.getProfile(refreshToken)

print("Got getProfile data at")
m=datetime.datetime.now()
print(m.second)


smartApi.generateToken(refreshToken)

print("Got generateToken at")
m=datetime.datetime.now()
print(m.second)


res = res['data']['exchanges']

print("Got res data at")
m=datetime.datetime.now()
print(m.second)

#
# #place order
# try:
#     orderparams = {
#         "variety": "NORMAL",
#         "tradingsymbol": "BANKNIFTY01NOV2342300PE",
#         "symboltoken": "47952",
#         "transactiontype": "BUY",
#         "exchange": "NFO",
#         "ordertype": "MARKET",
#         "producttype": "INTRADAY",
#         "duration": "DAY",
#         # "price": "150",
#         "squareoff": "0",
#         "stoploss": "0",
#         "quantity": "15"}
#
#     orderId=smartApi.placeOrder(orderparams)
#     print("The order id is: {}" .format(orderId))
#     # orderid2 = smartApi.cancelOrder(format(orderId), "NORMAL")
#     # print("The order id is: {}".format(orderid2))
# except Exception as e:
#     print("Order placement failed: {}".format(e.message))


buyOrder = {
    "variety": "NORMAL",
    "tradingsymbol": "BANKNIFTY22NOV2344300PE",
    "symboltoken": "42112",
    "transactiontype": "BUY",
    "exchange": "NFO",
    "ordertype": "MARKET",
    "producttype": "INTRADAY",
    "duration": "DAY",
    # "price": "150",
    "squareoff": "0",
    "stoploss": "0",
    "quantity": "60"
}
print("buy order params")
m=datetime.datetime.now()
print(m.second)
sellOrder = {
    "variety": "NORMAL",
    "tradingsymbol": "BANKNIFTY22NOV2344300PE",
    "symboltoken": "42112",
    "transactiontype": "SELL",
    "exchange": "NFO",
    "ordertype": "MARKET",
    "producttype": "INTRADAY",
    "duration": "DAY",
    # "price": "150",
    "squareoff": "0",
    "stoploss": "0",
    "quantity": "60"
}
# print("sell order params at")
# m=datetime.datetime.now()
# print(m.second)
# print(smartApi.orderBook['data'][-1]['orderstatus'] == 'complete')
# orderid2=smartApi.placeOrder(buyOrder)
# print("The order id is: {}".format(orderid2))
# orderBk=smartApi.orderBook()
# print(orderBk)
# print("order book ")
# m=datetime.datetime.now()
# print(m.second)
# if orderBk['data'][0]['status'] == 'complete':
#     print("WOrld")
#     print(orderBk['data'][0])
# print("order book last instance")
# m=datetime.datetime.now()
# print(m.second)
# dt = datetime.datetime.now()
# print(dt.minute, dt.microsecond/1000)
# print("Loop")
# while(1):
#     dt = datetime.datetime.now()
#     print(dt.minute, dt.microsecond / 1000)
#     if dt.minute==4 and (dt.microsecond/1000)>=100:
#         print("Execute")
#         break

# print(orderBook['data'][0]['status']=='rejected')
# print(orderBook['data'][-1]['status']=='complete')
#
# print("Hello")
# print(smartApi.orderBook()['data'])
# print(smartApi.orderBook['data'])
#
# orderBk=smartApi.orderBook()
# print(orderBk['data'][-1])
# print(orderBk['data'][-1]['orderstatus'])


# orderId = smartApi.placeOrder(buyOrder)
# print(orderId)
# i=10
# t = datetime.datetime.now()
# while(t.second!=0):
#     t = datetime.datetime.now()
#     print("Time before placing the order",t.minute,t.second,t.microsecond/1000)
#     print("i=",i)
#     i=i-1



# from here

#buy
while 1:
    t = datetime.datetime.now()
    if t.minute >= 15 and (t.microsecond/1000)>=25:
        t = datetime.datetime.now()
        print("Time before placing the order",t.minute,t.second,t.microsecond/1000)
        orderId = smartApi.placeOrder(buyOrder)
        t = datetime.datetime.now()
        print("Time after placing the order and getting the order id ",t.minute,t.second,t.microsecond/1000)
        print("The order id is: {}".format(orderId))
        break
    print(t.second)

# orderBook = smartApi.orderBook() #takes 5-6 sec to execute
# print("Order status before 10 second timer", orderBook['data'][-1]['orderstatus'])
t = datetime.datetime.now()
print("Time before 10 second sleep",t.minute,t.second,t.microsecond/1000)
print("starting 10 second sleep")
time.sleep(15)
print("ending 10 second sleep")
# orderBook = smartApi.orderBook()
# print("Order status before 10 second timer", orderBook['data'][-1]['orderstatus'])
t = datetime.datetime.now()
print("Time after 10 second sleep",t.minute,t.second,t.microsecond/1000)

# sell

while(1):
    try:
        orderBook = smartApi.orderBook()
        print(orderBook)
        if orderBook['data'][-1]['orderstatus'] == 'complete':
            t = datetime.datetime.now()
            print("Time before selling the order", t.minute, t.second, t.microsecond / 1000)
            orderid2 = smartApi.placeOrder(sellOrder)
            t = datetime.datetime.now()
            print("Time after selling the order", t.minute, t.second, t.microsecond / 1000)
            print("The order id is: {}".format(orderid2))
            break
    except Exception as e:
        print(e)

print("DONE")



# till here


# t1=time.time()
# t=datetime.datetime.now()
# print(t.hour, t.minute, t.second)
# t2=time.time()
# print(t2-t1)
#
# start_time = datetime.datetime.now()
#
# x = 0
# for i in range(1000):
#    x += i
#
# end_time = datetime.datetime.now()
#
# time_diff = (end_time - start_time)
# execution_time = time_diff.total_seconds() * 1000
#
# print(execution_time)
# smartApi.

# #gtt rule creation
#
# a=1
# while(1):
#     if (a >= 100):
#         try:
#             gttCreateParams = {
#                 "tradingsymbol": "SBIN-EQ",
#                 "symboltoken": "3045",
#                 "exchange": "NSE",
#                 "producttype": "MARGIN",
#                 "transactiontype": "BUY",
#                 "price": 100000,
#                 "qty": 11,
#                 "disclosedqty": 10,
#                 "triggerprice": 200000,
#                 "timeperiod": 365
#             }
#             rule_id = smartApi.gttCreateRule(gttCreateParams)
#             print("The GTT rule id is: {}".format(rule_id))
#             break
#         except Exception as e:
#             print("GTT Rule creation failed: {}".format(e.message))
#     else:
#         print(a)
#         a=a+1


#
# ltp=smartApi.ltpData("NSE", "SBIN-EQ", "3045")
# print("Ltp Data :", ltp)

# gtt rule list
# try:
#     status=["FORALL"] #should be a list
#     page=1
#     count=10
#     lists=smartApi.gttLists(status,page,count)
# except Exception as e:
#     print("GTT Rule List failed: {}".format(e.message))


# #Historic api
# try:
#     print(1)
#     historicParam={
#     "exchange": "NSE",
#     "symboltoken": "3045",
#     "interval": "ONE_MINUTE",
#     "fromdate": "2023-08-28 09:00",
#     "todate": "2023-08-29 09:16"
#     }
#     print(smartApi.getCandleData(historicParam))
# except Exception as e:
#     print("Historic Api failed: {}".format(e.message))
# logout
# try:
#     logout=smartApi.terminateSession('Your Client Id')
#     print("Logout Successfull")
# except Exception as e:
#     print("Logout failed: {}".format(e.message))
