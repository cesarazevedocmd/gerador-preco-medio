class MyAsset:
    asset = ""
    qty = 0
    price = 0

    def __init__(self, asset, qty, price):
        self.asset = asset
        self.qty = float(qty)
        self.price = float(price.replace(',', '.'))

    def print(self):
        print(self.asset + " QNT:" + str(self.qty) + " PRICE:" + str(self.price))