contract_prices = {
    'Contract1': [100.0, 110.0, 120.0, 130.0, 140.0],
    'Contract2': [150.0, 160.0, 170.0, 180.0, 190.0, 200],
    'Contract3': [120.5, 130.5, 140.5, 150.5, 160.5, 500, 600, 700],
    'Contract4': [200.0, 210.0, 220.0, 230.0],
    'Contract5': [180.0, 190.0, 200.0, 210.0, 220.0, 300, 400],
    'Contract6': [180.0, 190.0, 200.0],
    'Time': ["20:08:48"]
}

for key, value in contract_prices.items():
    print(key)
    print(value)
    print(value[0])