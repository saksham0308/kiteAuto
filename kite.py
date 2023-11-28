import datetime
from openpyxl import load_workbook
import time
from kiteconnect import KiteTicker
from selenium import webdriver
from kiteconnect import KiteConnect
from selenium.webdriver.common.by import By

my_file = open("kiteAuto/loginData.txt", "r")

# reading the file
data = my_file.read()
# replacing end splitting the text
# when newline ('\n') is seen.
myData = data.split("\n")
my_file.close()
kite = KiteConnect(api_key=myData[0])
# Redirect the user to the login url obtained
# from kite.login_url(), and receive the request_token
# from the registered redirect url after the login flow.
# Once you have the request_token, obtain the access_token
# as follows.

# This is not needed if chromedriver is already on your path:

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
driver.get(kite.login_url())
driver.implicitly_wait(10)

driver.find_element(By.ID, 'userid').send_keys(myData[1])
driver.find_element(By.ID, "password").send_keys(myData[2])
driver.find_element(By.CLASS_NAME, "button-orange").click()
time.sleep(20)

start = driver.current_url.find("request_token=") + 14
end = driver.current_url.find("&action")
if end == -1:
    myToken = driver.current_url[start:]
else:
    myToken = driver.current_url[start:end]
print(myToken)

data = kite.generate_session(myToken, api_secret=myData[3])
kite.set_access_token(data["access_token"])
print("Session started")

#
# # Place an order
# BuyOrder


# while 1:
#     t = datetime.datetime.now()
#     if t.minute >= 15 and (t.microsecond / 1000) >= 50:
#         t = datetime.datetime.now()
#         print("Time before placing the order", t.minute, t.second, t.microsecond / 1000)
#         try:
#             order_id = kite.place_order(tradingsymbol="BANKNIFTY23NOV43800PE",
#                                         exchange=kite.EXCHANGE_NFO,
#                                         transaction_type=kite.TRANSACTION_TYPE_BUY,
#                                         quantity=15,
#                                         variety=kite.VARIETY_REGULAR,
#                                         order_type=kite.ORDER_TYPE_MARKET,
#                                         product=kite.PRODUCT_MIS,
#                                         validity=kite.VALIDITY_DAY)
#
#             t = datetime.datetime.now()
#             print("Time after placing the order and getting the order id ", t.minute, t.second, t.microsecond / 1000)
#             print("The order id is: {}".format(order_id))
#         except Exception as e:
#             print("Order placement failed: {}".format(e))
#         print("Breaking buyOrder Loop")
#         break


kws = KiteTicker("i59md2vtn6mwc3kd", data["access_token"])
file_source = r'kiteAuto/BNData.xlsx'
a = 3
used = False
curRow = 0


def on_ticks(ws, ticks):
    global a
    global used
    global curRow

    # Callback to receive ticks.
    print("Ticks: {}".format(ticks))
    print(ticks[0]['last_price'])
    # print(ticks[1]['last_price'])
    now = datetime.datetime.now()
    contract_name = "ContractName"
    print(str(now.second) + " " + str(now.microsecond / 1000))
    workbook = load_workbook(filename=file_source)
    # Pick the sheet "new_sheet"
    ws4 = workbook["data"]
    format = '%Y-%m-%d %H:%M:%S'
    format2 = '%H:%M:%S'

    # applying strftime() to format the datetime
    dateTime = now.strftime(format)
    if not used:
        curRow = ws4.max_row + 1
        used = True

    ws4.cell(row=curRow, column=1).value = dateTime
    ws4.cell(row=curRow, column=2).value = contract_name

    secondIndex = 0
    try:
        secondIndex = ticks[1]['last_price']
    except Exception as E:
        print("Second Index not Available".format(E))
    # modify the cell
    cur_value = str(int(ticks[0]['last_price'])) + " " + now.strftime(format2) + " " + str(int(secondIndex))
    ws4.cell(row=curRow, column=a).value = cur_value

    # save the file
    workbook.save(filename=file_source)
    a += 1


def on_connect(ws, response):
    # Callback on successful connect.
    # Subscribe to a list of instrument_tokens (RELIANCE and ACC here).
    ws.subscribe([260105, 14919170])
    ws.set_mode(ws.MODE_FULL, [260105, 14919170])
    # 260105==banknifty


# Assign the callbacks.
kws.on_ticks = on_ticks
kws.on_connect = on_connect

# Infinite loop on the main thread. Nothing after this will run.
# You have to use the pre-defined callbacks to manage subscriptions.
kws.connect(threaded=True)

t = datetime.datetime.now()
print("Time before 10 second sleep", t.minute, t.second, t.microsecond / 1000)
print("Starting 10 second sleep")
time.sleep(10)
print("Ending 10 second sleep")
t = datetime.datetime.now()
print("Time after 10 second sleep", t.minute, t.second, t.microsecond / 1000)

# kite.cancel_order(variety=kite.VARIETY_AMO,
#                   order_id='231120000001027')

# orderBook = kite.orders()
# print(orderBook)
# SEll ORDER

# try:
#     orderBook = kite.orders()
#     print(orderBook)
#     print(orderBook[-1]['status'])
#     if orderBook[-1]['status'] == 'COMPLETE':
#         t = datetime.datetime.now()
#         print("Time before selling the order", t.minute, t.second, t.microsecond / 1000)
#         orderid2 = kite.place_order(tradingsymbol="BANKNIFTY23NOV43800PE",
#                                     exchange=kite.EXCHANGE_NFO,
#                                     transaction_type=kite.TRANSACTION_TYPE_SELL,
#                                     quantity=15,
#                                     variety=kite.VARIETY_REGULAR,
#                                     order_type=kite.ORDER_TYPE_MARKET,
#                                     product=kite.PRODUCT_MIS,
#                                     validity=kite.VALIDITY_DAY)
#         t = datetime.datetime.now()
#         print("Time after selling the order", t.minute, t.second, t.microsecond / 1000)
#         print("The order id is: {}".format(orderid2))
# except Exception as e:
#     print(e)

print("DONE Buy/Sell")
print("Starting extra 5 second sleep")
time.sleep(5)
print("Ending extra 5 second sleep")
print("Session Ended")

# Get instruments
# with open('readme.txt', 'w') as f:
#     f.write(str(kite.instruments('NFO')))
