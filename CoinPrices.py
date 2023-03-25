import openpyxl
from bs4 import BeautifulSoup
import requests

def updateValues():
    path = "C:/Users/jorgl/OneDrive/Escritorio/Guide.xlsx"

    for i in range(1):
        try:
            cryptos = ["bitcoin", "ethereum", "bnb", "cardano", "solana", "polkadot-new", "monero",
                       "basic-attention-token", "harmony"]

            cryptoExcel = ["Bitcoin (BTC)", "Ethereum (ETH)", "Binance Coin (BNB)", "Cardano (ADA)", "Solana (SOL)",
                           "Polkadot (DOT)", "Monero (XMR)", "Brave Bat", "Harmony ONE"]

            agent = {"User-Agent": "Brave"}
            priceList = []

            for item in cryptos:
                site = "https://coinmarketcap.com/currencies/{}/".format(item)
                request = requests.get(site, headers=agent)
                b = BeautifulSoup(request.text, "html.parser")

                price = float(b.find("div", attrs={'class': 'priceValue'}).text.removeprefix("$").replace(",", ""))
                priceList.append(price)

            wb = openpyxl.load_workbook(path, read_only=False, keep_vba=False)
            sheet1 = wb.active

            index = 0
            for column in sheet1.iter_cols(min_col=1, max_col=1, min_row=2, max_row=10):
                for cell in column:
                    cell.value = cryptoExcel[index]
                    index += 1

            index = 0
            for column in sheet1.iter_cols(min_col= 2, max_col= 2, min_row= 2, max_row= 10):
                for cell in column:
                    cell.value = priceList[index]
                    index += 1
            wb.save(path)
            print("Excel file saved with success\n")

        except Exception as ex:
            print(ex)


if __name__ == '__main__':
    try:
        updateValues()
    except KeyboardInterrupt:
        print("Program terminated by user")