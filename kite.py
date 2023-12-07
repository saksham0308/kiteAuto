import datetime
from openpyxl import load_workbook
import time
from kiteconnect import KiteTicker
from selenium import webdriver
from kiteconnect import KiteConnect
from selenium.webdriver.common.by import By
import pandas as pd
import json

# my_file = open("loginData.json", "r")
my_file = "loginData.json"
with open(my_file, "r") as file:
    myData = json.load(file)
kite = KiteConnect(api_key=myData.get("api_key"))
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

driver.find_element(By.ID, 'userid').send_keys(myData.get("userid"))
driver.find_element(By.ID, "password").send_keys(myData.get("password"))
driver.find_element(By.CLASS_NAME, "button-orange").click()
time.sleep(20)

start = driver.current_url.find("request_token=") + 14
end = driver.current_url.find("&action")
if end == -1:
    myToken = driver.current_url[start:]
else:
    myToken = driver.current_url[start:end]
print(myToken)

data = kite.generate_session(myToken, api_secret=myData.get("api_secret"))
kite.set_access_token(data["access_token"])
print("Session started")

kws = KiteTicker(myData.get("api_key"), data["access_token"])
file_source = r'BNData.xlsx'
a = 3
used = False
curRow = 0
df = 0
df_transposed = 0
sheet_name = 0
last_price = 0
flag = False
strike_price = 0
option_strike = ""
prev1Token = 0
prev2Token = 0
next1Token = 0
next2Token = 0


def round_to_nearest_hundred(number):
    remainder = number % 100

    # Determine whether to round up or down based on the remainder
    if remainder >= 50:
        rounded_number = number + (100 - remainder)
    else:
        rounded_number = number - remainder

    return rounded_number


def create_option_strike(symbol, date, rounded_number, option_type='PE'):
    return f"{symbol}{date}{rounded_number:05d}{option_type}"


contract_prices = {
    'Time': [],
    '260105': ["BANKNIFTY"]
}


def on_ticks(ws, ticks):
    global a, prev1Token, prev2Token, next1Token, next2Token
    global used
    global curRow
    global df
    global df_transposed
    global contract_prices
    global sheet_name
    global last_price
    global flag
    global strike_price
    global option_strike

    if not flag:
        t = datetime.datetime.now()
        print("time before flag false ", t.minute, t.second, t.microsecond / 1000)
        if (ticks[0]['last_price']) != (myData.get("bankNiftyLastPrice")):
            t = datetime.datetime.now()
            print("time after flag true ", t.minute, t.second, t.microsecond / 1000)
            # original_number = 45678
            symbol = myData.get("symbol")
            date = myData.get("date")
            # original_number = 45678
            strike_price = round_to_nearest_hundred(int(ticks[0]['last_price']))
            option_strike = create_option_strike(symbol, date, strike_price)
            # get subscribe token from readme for strikePrice and 4 other prices upper and lower
            flag = True
            print(flag)
            strikePricePrev1 = strike_price - 200
            strikePricePrev2 = strike_price - 100
            strikePriceNext1 = strike_price + 100
            strikePriceNext2 = strike_price + 200
            # print(strikePricePrev1)
            # print(strikePricePrev2)
            # print(strikePriceNext1)
            # print(strikePriceNext2)

            option_strikePrev1 = create_option_strike(symbol, date, strikePricePrev1)
            option_strikePrev2 = create_option_strike(symbol, date, strikePricePrev2)
            option_strikeNext1 = create_option_strike(symbol, date, strikePriceNext1)
            option_strikeNext2 = create_option_strike(symbol, date, strikePriceNext2)

            # print(option_strikePrev1)
            # print(option_strikePrev2)
            # print(option_strikeNext1)
            # print(option_strikeNext2)

            prev1Token = myData.get(str(option_strikePrev1))
            prev2Token = myData.get(str(option_strikePrev2))
            optionStrikeToken = myData.get(str(option_strike))
            next1Token = myData.get(str(option_strikeNext1))
            next2Token = myData.get(str(option_strikeNext2))

            # print(prev1Token)
            # print(prev2Token)
            # print(next1Token)
            # print(next2Token)

            ws.subscribe([int(myData.get("BANKNIFTYTOKEN")),prev1Token, prev2Token, optionStrikeToken, next1Token, next2Token])
            ws.set_mode(ws.MODE_FULL, [int(myData.get("BANKNIFTYTOKEN")),prev1Token, prev2Token, optionStrikeToken, next1Token, next2Token])
            contract_prices[str(prev1Token)] = [str(option_strikePrev1)]
            contract_prices[str(prev2Token)] = [str(option_strikePrev2)]
            contract_prices[str(optionStrikeToken)] = [str(option_strike)]
            contract_prices[str(next1Token)] = [str(option_strikeNext1)]
            contract_prices[str(next2Token)] = [str(option_strikeNext2)]

    # Callback to receive ticks.
    print(ticks)
    now = datetime.datetime.now()
    # book = load_workbook(filename=file_source)
    format2 = '%H:%M:%S'

    # save the file
    # book.save(filename=file_source)
    # a += 1
    # Example dictionary with contract names as keys and prices as values
    # contract_prices['Time'].append(now.strftime(format2))
    contract_prices['Time'].append(now.strftime(format2))
    for key, value in contract_prices.items():
        if str(key) == 'Time':
            continue
        check = 1
        for tick in ticks:
            # print(str(key) + " "+str(tick['instrument_token']))
            if str(key) == str(tick['instrument_token']):
                contract_prices[key].append(tick['last_price'])
                # print(str(key) + " " + str(tick['instrument_token']))
                # print(contract_prices)
                check = 0
                break
        if check == 1:
            contract_prices[key].append(0)


# , 10483458,10490114,10498562,10499074,10499586

def on_connect(ws, response):
    # Callback on successful connect.
    # Subscribe to a list of instrument_tokens (RELIANCE and ACC here).
    ws.subscribe([myData.get("BANKNIFTYTOKEN")])
    ws.set_mode(ws.MODE_FULL, [myData.get("BANKNIFTYTOKEN")])
    # 260105==banknifty


# Assign the callbacks.
kws.on_ticks = on_ticks
kws.on_connect = on_connect

# Infinite loop on the main thread. Nothing after this will run.
# You have to use the pre-defined callbacks to manage subscriptions.
kws.connect(threaded=True)
# Place an order
# BuyOrder

while 1:
    t = datetime.datetime.now()
    if t.minute >= int(myData.get("timeInMinutes")) and (t.microsecond / 1000) >= 50 and flag:
        print("Time before placing the order", t.minute, t.second, t.microsecond / 1000)
        try:
            order_id = kite.place_order(tradingsymbol=option_strike,
                                        exchange=kite.EXCHANGE_NFO,
                                        transaction_type=kite.TRANSACTION_TYPE_BUY,
                                        quantity=int(myData.get("quantity")),
                                        variety=kite.VARIETY_REGULAR,
                                        order_type=kite.ORDER_TYPE_MARKET,
                                        product=kite.PRODUCT_MIS,
                                        validity=kite.VALIDITY_DAY)

            t = datetime.datetime.now()
            print("Time after placing the order and getting the order id ", t.minute, t.second, t.microsecond / 1000)
            print("The order id is: {}".format(order_id))
        except Exception as e:
            print("Order placement failed: {}".format(e))
        print("Breaking buyOrder Loop")
        break


t = datetime.datetime.now()
print("Time before 10 second sleep", t.minute, t.second, t.microsecond / 1000)
time.sleep(9)
t = datetime.datetime.now()
print("Time after 10 second sleep", t.minute, t.second, t.microsecond / 1000)

# kite.cancel_order(variety=kite.VARIETY_AMO,
#                   order_id='231120000001027')

orderBook = kite.orders()
print(orderBook)
# SEll ORDER

try:
    orderBook = kite.orders()
    print(orderBook)
    print(orderBook[-1]['status'])
    if orderBook[-1]['status'] == 'COMPLETE':
        t = datetime.datetime.now()
        print("Time before selling the order", t.minute, t.second, t.microsecond / 1000)
        orderid2 = kite.place_order(tradingsymbol=option_strike,
                                    exchange=kite.EXCHANGE_NFO,
                                    transaction_type=kite.TRANSACTION_TYPE_SELL,
                                    quantity=int(myData.get("quantity")),
                                    variety=kite.VARIETY_REGULAR,
                                    order_type=kite.ORDER_TYPE_MARKET,
                                    product=kite.PRODUCT_MIS,
                                    validity=kite.VALIDITY_DAY)
        t = datetime.datetime.now()
        print("Time after selling the order", t.minute, t.second, t.microsecond / 1000)
        print("The order id is: {}".format(orderid2))
except Exception as e:
    print(e)

print("DONE Buy/Sell")
print("Starting extra 5 second sleep")
time.sleep(5)
print("Session Ended")


def append_df_to_excel(df, excel_path):
    try:
        # Load the existing Excel workbook
        book = load_workbook(excel_path)
        currentMonth = str(t.today().strftime("%B"))
        # Get the existing sheet or create a new one
        sheet = book[currentMonth] if currentMonth in book.sheetnames else book.create_sheet(title=currentMonth)

        # Find the last used row dynamically
        last_used_row = sheet.max_row if sheet.max_row > 1 else 1

        for key in df.index:
            last_used_row += 1
            sheet.cell(row=last_used_row, column=1, value=key)

            # Write values from the DataFrame to subsequent columns
            for idx, value in enumerate(df.loc[key].astype(str).values, start=2):
                sheet.cell(row=last_used_row, column=idx, value=value)

        # Save the changes to the existing Excel file
        book.save(excel_path)
        print(f"DataFrame appended to {excel_path} under {sheet_name} starting from row {last_used_row}")

    except Exception as e:
        print(f"Error appending DataFrame to {excel_path}: {e}")


dynamic_range = max(len(values) if isinstance(values, list) else 0 for values in contract_prices.values())
df = pd.DataFrame.from_dict(contract_prices, orient='index')

# Display the DataFrame
print(df)
append_df_to_excel(df, 'BNData.xlsx')

# Get instruments
# with open('readme.txt', 'w') as f:
#     f.write(str(kite.instruments('NFO')))
#
