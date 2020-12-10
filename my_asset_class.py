class MyAsset:
    asset = ""
    quant = 0
    price = 0

    def __init__(self, asset, quant, price):
        self.asset = asset
        self.quant = float(quant)
        self.price = float(price.replace(',', '.'))

    def print(self):
        print(self.asset + " QNT:" + str(self.quant) + " PRICE:" + str(self.price))