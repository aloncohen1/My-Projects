# This code will automatically update your sales on eBay to a Google sheet

# Preparations:
# Please prepare credentials to eBay's API and Google sheets API
# Google sheet API: https://developers.google.com/sheets/api/
# eBay API: https://developer.ebay.com/tools/quick-start
# Open a new google sheet and copy the sheet's ID (sheet's name should be "Sheet1") 

# Notice: the orders will be updated in the sheet only after full payment has been received

# Made by: Alon Cohen, alon.cohen8@mail.huji.ac.il

# imports
from ebaysdk.trading import Connection as Trading
from ebaysdk.exception import ConnectionError

import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from pprint import pprint
from googleapiclient import discovery

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


# connect to your google API


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'enter your CLIENT SECRET FILE here (jason file)'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = None
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')

credentials = None

service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

# The ID of the spreadsheet to update.

spreadsheet_id = "Enter you spreadsheet id here can be found in the sheet's URL"  # TODO: Update placeholder value.

# The A1 notation of the values to update.
range_ = "Sheet1!A1:N57"  # TODO: Update placeholder value.

# How the input data should be interpreted.
value_input_option = 'USER_ENTERED'  # TODO: Update placeholder value.
result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_).execute()
values = result.get('values', [])

# Creates a title for the columns
range_ = "Sheet1!A1:V57"
value_range_body = {"range": "Sheet1!A1:V57",
                                        "majorDimension": "ROWS",
                                        "values": [['Imported data', 'Imported data','Imported data', 'Imported data', 'Imported data', 'Imported data',
                                                    'Excel function','Insert manually','Excel function','Excel function','Imported data','Insert manually',
                                                    'Insert manually','Insert manually','Imported data','Imported data','Insert manually'], ], }
request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                                     valueInputOption=value_input_option, body=value_range_body)


response = request.execute()
pprint(response)

range_ = "Sheet1!A2:V57"
value_range_body = {"range": "Sheet1!A2:V57",
                                        "majorDimension": "ROWS",
                                        "values": [['Product', 'Quantity','Product type', 'Country', 'Date', 'Price','Income after fees','Shipping cost',
                                                    'Cost of the product','Net revenue','Selling Platform','Package shipped?','Tacking mumber/ No tacking mumber',
                                                    'tacking mumber','User ID','Email','Notes'], ], }
request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                                     valueInputOption=value_input_option, body=value_range_body)

response = request.execute()
pprint(response)


range_ = "Sheet1!A3:V3"  # TODO: Update placeholder value.

# How the input data should be interpreted.
value_input_option = 'USER_ENTERED'  # TODO: Update placeholder value.
result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_).execute()
test = result.get('values', [])

# A test row- in case there are no sales in the sheet
if test==[]:
    range_ = "Sheet1!A3:V3"
    value_range_body = {"range": "Sheet1!A3:V3",
                        "majorDimension": "ROWS",
                        "values": [['Product A', '3', '', 'US', '2014-08-02T07:16:07.000Z',
                                    '', '', '', '',
                                    '', '', '', '',
                                    '', '', '', ''], ], }
    request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                     valueInputOption=value_input_option, body=value_range_body)

    response = request.execute()
    pprint(response)
range_ = "Sheet1!A1:N57"  # TODO: Update placeholder value.

# How the input data should be interpreted.
value_input_option = 'USER_ENTERED'  # TODO: Update placeholder value.
result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_).execute()
values = result.get('values', [])

# count how many rows with values there are in the sheet file.
excel_row_count=1
for row in values:
    if row !=[]:
        excel_row_count+=1

# creat a list of all of the dates
excel_all_dates_list=[]
for row in values:
    excel_all_dates_list.append(row[4])

excel_all_date=excel_all_dates_list[-1] # the selling date of the last product I sold according to the sheet file
excel_date=excel_all_date.split('T')[0] # only date (no hours/minutes)
excel_time=excel_all_date.split('T')[1] # hours and minutes
excel_year=excel_all_date.split('T')[0].split('-')[0] # year
excel_month=excel_all_date.split('T')[0].split('-')[1] # month
excel_day=excel_all_date.split('T')[0].split('-')[2] # day
excel_hour=excel_all_date.split('T')[1].split(':')[0] # hour
excel_min=excel_all_date.split('T')[1].split(':')[1] # minute

# accessing to the ebay API
try:
    api = Trading(appid="enter your appid here", devid="enter your devid here",
                  certid="enter your certid here",
                  token="enter your token here",
                  config_file=None)
    soldlists = api.execute('GetMyeBaySelling', {'SoldList': True,
                                                 'DetailLevel': 'ReturnAll',
                                                 'PageNumber': 0})
    response_ebay = api.execute('GetUser', {})
    ebay_api = soldlists.dict()
    allsells = ebay_api['SoldList']['OrderTransactionArray']['OrderTransaction']
    allsells_count = -1

    row_x = excel_row_count
    row_y = excel_row_count + 5

    # Defines the selling date of each item from the eBay API
    for i in allsells:
        try:
            if allsells[allsells_count]['Transaction']['SellerPaidStatus'] == 'NotPaid':
                allsells_count -= 1
            if allsells[allsells_count]['Transaction']['SellerPaidStatus'] == 'Refunded':
                allsells_count -= 1
            if allsells[allsells_count]['Transaction']['SellerPaidStatus'] == 'PaymentPending':
                allsells_count -= 1
            ebay_all_date = allsells[allsells_count]['Transaction']['PaidTime']
            ebay_date = ebay_all_date.split('T')[0]
            ebay_time = ebay_all_date.split('T')[1]
            ebay_year = ebay_all_date.split('T')[0].split('-')[0]
            ebay_month = ebay_all_date.split('T')[0].split('-')[1]
            ebay_day = ebay_all_date.split('T')[0].split('-')[2]
            ebay_hour = ebay_all_date.split('T')[1].split(':')[0]
            ebay_min = ebay_all_date.split('T')[1].split(':')[1]

            # checking if the item is new for the excel sheet (by using his date)
            if (excel_year > ebay_year or (excel_year == ebay_year and excel_month > ebay_month) or (
                                excel_year == ebay_year and excel_month == ebay_month and excel_day > ebay_day) or (
                                    excel_year == ebay_year and excel_month == ebay_month and
                                excel_day == ebay_day and excel_hour > ebay_hour) or (
                                        excel_year == ebay_year and excel_month == ebay_month and
                                    excel_day == ebay_day and excel_hour == ebay_hour and excel_min > ebay_min) or (
                            excel_date == ebay_date and excel_time == ebay_time)):
                allsells_count -= 1
                continue

            else:
                # Defines the selling data of each item from the eBay API
                Item_name = allsells[allsells_count]['Transaction']['Item']['Title']  # Item Name
                Quantity = allsells[allsells_count]['Transaction']['QuantityPurchased']  # Quantity
                Price = allsells[allsells_count]['Transaction']['TotalTransactionPrice']['value']  # Price
                Country = allsells[allsells_count]['Transaction']['Buyer']['BuyerInfo']['ShippingAddress'][
                    'Country']  # Country
                Paid_Date = allsells[allsells_count]['Transaction']['PaidTime']  # Paid Date
                Email = allsells[allsells_count]['Transaction']['Buyer']['Email']  # Email
                Buyer_ID = allsells[allsells_count]['Transaction']['Buyer']['UserID']  # User
                Item_shade = ''

                # separating added information, like 'product type', if exist
                if len(Item_name.split('[')) > 1:
                    Item_shade = str(Item_name.split('[')[1].split(']')[0])
                    Item_name = str(Item_name.split('[')[0])

                row_x = str(row_x)
                row_y = str(row_y)

                # Excel functions that calculate costs, tax and net revenue
                item_cost_function = '=B' + row_x + '*((if(A' + row_x + '="name of profuct A",100,0) + if(A' + row_x + '="name of profuct B",70,0)))'  # and so on...

                # Relevant to Israeli sellers
                tax_function = '=F' + row_x + '-(F' + row_x + '*0.1)-0.3-if(or(D' + row_x + '="US",D' + row_x + '="CA",D' + row_x + '="PR",D' + row_x + '="GU"),F' + row_x + '*0.044,0)-if(or(D' \
                               + row_x + '="DK",D' + row_x + '="FI",D' + row_x + '="GL",D' + row_x + '="IS",D' + row_x + '="NO",D' + row_x + '="SE"),F' + row_x + \
                               '*0.038,0)-if(or(D' + row_x + '="AT",D' + row_x + '="BE",D' + row_x + '="CY",D' + row_x + '="EE",D' + row_x + '="FR",D' + row_x + \
                               '="DE",D' + row_x + '="GI",D' + row_x + '="GR",D' + row_x + '="IE",D' + row_x + '="IT",D' + row_x + '="LU",D' + row_x + '="MT",D' + \
                               row_x + '="MC",D' + row_x + '="ME",D' + row_x + '="NL",D' + row_x + '="PT",D' + row_x + '="SM",D' + row_x + '="SK",D' + row_x + '="SI",D' \
                               + row_x + '="GB"),F' + row_x + '*0.039,0)-if(or(D' + row_x + '="CZ",D' + row_x + '="UK",D' + row_x + '="LT",D' + row_x + '="Kosovo",D' + \
                               row_x + '="MD",D' + row_x + '="PL",D' + row_x + '="CH",D' + row_x + '="RO",D' + row_x + '="RU",D' + row_x + '="UA"),F' + row_x + \
                               '*0.047,0)-if(or(D' + row_x + '="AU",D' + row_x + '="JP",D' + row_x + '="IL",D' + row_x + '="MV",D' + row_x + '="PE",D' + row_x + '="BR",D' \
                               + row_x + '="TH",D' + row_x + '="CO",D' + row_x + '="NZ",D' + row_x + '="EC",D' + row_x + '="ES",D' + row_x + '="MU",D' + row_x + '="SG",D' + \
                               row_x + '="CL",D' + row_x + '="HK",D' + row_x + '="MA",D' + row_x + '="HR",D' + row_x + '="HN",D' + row_x + '="LK",D' + row_x + '="UY",D' + row_x + \
                               '="KR",D' + row_x + '="MK",D' + row_x + '="AZ",D' + row_x + '="CN",D' + row_x + '="IN",D' + row_x + '="SV",D' + row_x + '="NG",D' + row_x + '="MX",D' + row_x + \
                               '="ZA",D' + row_x + '="AM",D' + row_x + '="RS",D' + row_x + '="AR",D' + row_x + '="HU"),F' + row_x + '*0.054,0)'

                net_rev = '=G' + row_x + '*$F$71-H' + row_x + '-I' + row_x

                # write the values into the Google spreadsheet
                range_ = "Sheet1!A" + row_x + ":N" + row_y
                value_range_body = {"range": "Sheet1!A" + row_x + ":N" + row_y,
                                    "majorDimension": "ROWS",
                                    "values": [
                                        [Item_name, Quantity, Item_shade, Country, Paid_Date, Price, tax_function, '',
                                         item_cost_function, net_rev, 'eBay'], ], }
                request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                                 valueInputOption=value_input_option,
                                                                 body=value_range_body)
                response = request.execute()
                range_ = "Sheet1!O" + row_x + ":R" + row_y
                value_range_body = {"range": "Sheet1!O" + row_x + ":R" + row_y,
                                    "majorDimension": "ROWS",
                                    "values": [[Buyer_ID, Email, ], ], }
                request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                                 valueInputOption=value_input_option,
                                                                 body=value_range_body)
                response = request.execute()

                row_x = int(row_x)
                row_y = int(row_y)
                row_x += 1
                row_y += 1

                # TODO: Change code below to process the `response` dict:
                pprint(response)

                allsells_count -= 1

                # this part is similar to the previous one. it's a little different for cases in which a buyer purchase more than one product.
        except:
            for i in allsells[allsells_count]['Order']['TransactionArray']['Transaction']:
                try:
                    ebay_all_date = i['PaidTime']
                    ebay_date = ebay_all_date.split('T')[0]
                    ebay_time = ebay_all_date.split('T')[1]
                    ebay_year = ebay_all_date.split('T')[0].split('-')[0]
                    ebay_month = ebay_all_date.split('T')[0].split('-')[1]
                    ebay_day = ebay_all_date.split('T')[0].split('-')[2]
                    ebay_hour = ebay_all_date.split('T')[1].split(':')[0]
                    ebay_min = ebay_all_date.split('T')[1].split(':')[1]

                    if (excel_year > ebay_year or (excel_year == ebay_year and excel_month > ebay_month) or (
                                        excel_year == ebay_year and excel_month == ebay_month and excel_day > ebay_day) or (
                                            excel_year == ebay_year and excel_month == ebay_month and
                                        excel_day == ebay_day and excel_hour > ebay_hour) or (
                                                excel_year == ebay_year and excel_month == ebay_month and
                                            excel_day == ebay_day and excel_hour == ebay_hour and excel_min > ebay_min) or (
                                    excel_date == ebay_date and excel_time == ebay_time)):
                        allsells_count -= 1

                        continue
                    else:
                        Item_name = i['Item']['Title']  # Item Name
                        Quantity = i['QuantityPurchased']  # Quantity
                        Price = float(i['TotalTransactionPrice']['value']) + float(
                            i['Item']['ShippingDetails']['ShippingServiceOptions']['ShippingServiceCost'][
                                'value'])  # Price
                        Price = str(Price)
                        Country = i['Buyer']['BuyerInfo']['ShippingAddress']['Country']  # Country
                        Paid_Date = i['PaidTime']  # Paid Date
                        Email = i['Buyer']['Email']  # Email
                        Buyer_ID = i['Buyer']['UserID']  # Buyer ID
                        Item_shade = ''

                        if len(Item_name.split('[')) > 1:
                            Item_shade = str(Item_name.split('[')[1].split(']')[0])
                            Item_name = str(Item_name.split('[')[0])

                        row_x = str(row_x)
                        row_y = str(row_y)

                        item_cost_function = '=B' + row_x + '*((if(A' + row_x + '="name of profuct A",100,0) + if(A' + row_x + '="name of profuct B",70,0)))'  # and so on...

                        tax_function = '=F' + row_x + '-(F' + row_x + '*0.1)-0.3-if(or(D' + row_x + '="US",D' + row_x + '="CA",D' + row_x + '="PR",D' + row_x + '="GU"),F' + row_x + '*0.044,0)-if(or(D' + row_x + '="DK",D' + row_x + '="FI",D' + row_x + '="GL",D' \
                                       + row_x + '="IS",D' + row_x + '="NO",D' + row_x + '="SE"),F' + row_x + '*0.038,0)-if(or(D' + row_x + '="AT",D' + row_x + '="BE",D' + row_x + '="CY",D' + row_x + '="EE",D' + row_x + '="FR",D' + row_x \
                                       + '="DE",D' + row_x + '="GI",D' + row_x + '="GR",D' + row_x + '="IE",D' + row_x + '="IT",D' + row_x + '="LU",D' + row_x + '="MT",D' + row_x + '="MC",D' + row_x + '="ME",D' + row_x + '="NL",D' + row_x \
                                       + '="PT",D' + row_x + '="SM",D' + row_x + '="SK",D' + row_x + '="SI",D' + row_x + '="GB"),F' + row_x + '*0.039,0)-if(or(D' + row_x + '="CZ",D' + row_x + '="UK",D' + row_x + '="LT",D' + row_x + '="Kosovo",D' \
                                       + row_x + '="MD",D' + row_x + '="PL",D' + row_x + '="CH",D' + row_x + '="RO",D' + row_x + '="RU",D' + row_x + '="UA"),F' + row_x + '*0.047,0)-if(or(D' + row_x + '="AU",D' + row_x + '="JP",D' + row_x + '="IL",D' \
                                       + row_x + '="MV",D' + row_x + '="PE",D' + row_x + '="BR",D' + row_x + '="TH",D' + row_x + '="CO",D' + row_x + '="NZ",D' + row_x + '="EC",D' + row_x + '="ES",D' + row_x + '="MU",D' + row_x + '="SG",D' + row_x \
                                       + '="CL",D' + row_x + '="HK",D' + row_x + '="MA",D' + row_x + '="HR",D' + row_x + '="HN",D' + row_x + '="LK",D' + row_x + '="UY",D' + row_x + '="KR",D' + row_x + '="MK",D' + row_x + '="AZ",D' + row_x + '="CN",D' \
                                       + row_x + '="IN",D' + row_x + '="SV",D' + row_x + '="NG",D' + row_x + '="MX",D' + row_x + '="ZA",D' + row_x + '="AM",D' + row_x + '="RS",D' + row_x + '="AR",D' + row_x + '="HU"),F' + row_x + '*0.054,0)'

                        net_rev = '=G' + row_x + '*$F$71-H' + row_x + '-I' + row_x

                        range_ = "Sheet1!A" + row_x + ":N" + row_y
                        value_range_body = {"range": "Sheet1!A" + row_x + ":N" + row_y,
                                            "majorDimension": "ROWS",
                                            "values": [[Item_name, Quantity, Item_shade, Country, Paid_Date, Price,
                                                        tax_function, '', item_cost_function, net_rev, 'eBay'], ], }
                        request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                                         valueInputOption=value_input_option,
                                                                         body=value_range_body)
                        response = request.execute()
                        range_ = "Sheet1!O" + row_x + ":R" + row_y
                        value_range_body = {"range": "Sheet1!O" + row_x + ":R" + row_y,
                                            "majorDimension": "ROWS",
                                            "values": [[Buyer_ID, Email, ], ], }
                        request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_,
                                                                         valueInputOption=value_input_option,
                                                                         body=value_range_body)
                        response = request.execute()
                        row_x = int(row_x)
                        row_y = int(row_y)
                        row_x += 1
                        row_y += 1

                        # TODO: Change code below to process the `response` dict:
                        pprint(response)
                    allsells_count -= 1
                except:
                    continue



except ConnectionError as e:
    print(e)
    print(e.response_ebay.dict())
