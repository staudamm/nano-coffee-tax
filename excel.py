HEADER_ROW = 14

AMOUNT = {'EU': 'C9',
          'Ausfuhr': 'C10'}

TIME_FROM = 'D6'
TIME_TO = 'D7'
TRACKER_COL = 'M'
ORDER_COL = 'R'

row_key_to_json_key = {
    "Amount": 'total_coffee_weight_sold#total_coffee_weight_sold',
    "Country": 'shipping_address.country',
    "Name": 'shipping_address.name',
    "Address": 'shipping_address.address1',
    "Zip": 'shipping_address.zip',
    "City": 'shipping_address.city',
    "Order ID": 'order_name'
}


row = {
    "ID": 1,
    "Sender": "NANO Kaffee GmbH, Charlottenstr. 1, 10969, Berlin",
    "Place and Data of Export": "Berlin, XXXX",
    "Amount": 0,
    "Product": "100% Röstkaffee",
    "Region": "",
    "Country": "",
    "Name": "",
    "Address": "",
    "Zip": "",
    "City": "",
    "VAT Number": "",
    "Tracking URL": "",
    "Reference": "",
    "MRN": "",
    "Send Method": "",
    "Confirmation": "",
    "Order ID": ""
}
TEMPLATE_EXCEL_FILE = "NANO_KAFFEE_GmbH_YYYY_MM_Abt.3A.xlsx"
