import requests
import ezodf

apikey_file = open("/home/hugo/Perso/misc/financial_market/marketstack_apikey.txt")
apikey = apikey_file.read()[:-1]

excel_marks_file_path = "/home/hugo/Perso/banque/investissements.ods"
symbol_column = "O"
quote_column = "I"

api_url = "http://api.marketstack.com/v1/eod/latest?access_key=" + apikey


def get_quote_for_symbol(symbol):
    request_symbol = api_url + "&symbols=" + symbol
    r = requests.get(request_symbol)
    r_json = r.json()
    return r_json["data"][0]["close"]


doc = ezodf.opendoc(excel_marks_file_path)


TAA = doc.sheets["TAA"]
for cell_row in range(2, len(TAA.column(0)) + 1):
    symbol = TAA[symbol_column + str(cell_row)].value
    if symbol is not None:
        quote = get_quote_for_symbol(symbol)
        TAA[quote_column + str(cell_row)].set_value(quote)

doc.save()
