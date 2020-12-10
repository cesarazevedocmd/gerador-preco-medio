class Operation:
    asset = ""
    quant = 0
    price = 0
    type = ""

    def __init__(self, asset, quant, price, type):
        self.asset = asset
        self.quant = float(quant)
        self.price = float(price.replace(',', '.'))
        self.type = type

    def print(self):
        print(self.asset + " " + str(self.quant) + " " + str(self.price) + " | " + self.type)