import datetime
import math
import os

import requests
import xlsxwriter
from multiprocessing.dummy import Pool as ThreadPool

session = requests.session()
pool = ThreadPool(10)
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:64.0) Gecko/20100101 Firefox/64.0",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Accept-Encoding": "gzip, deflate, br",
    "Host": "www.seabourn.com",
    "Refer": "https://www.seabourn.com/en_US/find-a-cruise.html",
    "Origin": "https://www.seabourn.com",
    "country": "US",
    "Cookie": "countryCode=US; continentCode=NA; akaas_SBN_PROD=2147483647~rv=3~id=7865248ae1be4ccc34165c0c4e6ea975; languageCode=en_US; ak_bmsc=C6B83E26BF582C2516CB77531E6CAB8102142D6763300000994ACB5B9C784428~pla5fA63jw498/B1OLgl+p96XmG7dvXB7qBhqOzwrMmGYcQhkP2ak+lb3irs8AfzguX/+IHOKckan4JPcZZ7Bh6oYB82FayTdjUVIyQcmT1wEEND/e10ZrQvceyxyQp4hhGZhSa7Qd0ip0Js0bU2nxYkwuTgELHnKVNPm3Vjpz7Kq7EXpSjjZvNP4HOMFz+Xnbz4d8abCtkOcX12OSFK4NbU5ZWCeQrUQroA5D1ExF5Bc=",
    "currencycode": "USD",
    "loyaltynumber": "",
    "TE": "Trailers"
}
price_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:64.0) Gecko/20100101 Firefox/64.0",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Accept-Encoding": "gzip, deflate, br",
    "Host": "www.seabourn.com",
    "Refer": "https://www.seabourn.com/en_US/find-a-cruise.html",
    "Origin": "https://www.seabourn.com",
    "country": "US",
    "Cookie": "countryCode=US; continentCode=NA; akaas_SBN_PROD=2147483647~rv=3~id=7865248ae1be4ccc34165c0c4e6ea975; languageCode=en_US; ak_bmsc=C6B83E26BF582C2516CB77531E6CAB8102142D6763300000994ACB5B9C784428~pla5fA63jw498/B1OLgl+p96XmG7dvXB7qBhqOzwrMmGYcQhkP2ak+lb3irs8AfzguX/+IHOKckan4JPcZZ7Bh6oYB82FayTdjUVIyQcmT1wEEND/e10ZrQvceyxyQp4hhGZhSa7Qd0ip0Js0bU2nxYkwuTgELHnKVNPm3Vjpz7Kq7EXpSjjZvNP4HOMFz+Xnbz4d8abCtkOcX12OSFK4NbU5ZWCeQrUQroA5D1ExF5Bc=",
    "currencycode": "USD",
    "loyaltynumber": "",
    "TE": "Trailers",
    "brand": "sbn",
    "locale": "en_US"
}

urls = []
voyage_ids = []
voyages = []
to_write = []

codes = ['A', 'SN', 'I', 'O', 'P', 'N', 'C', 'Q', 'J', 'X', 'EM', 'EN', 'L', 'T', 'S', 'ET', 'W']

for code in codes:
    print("Downloading", "'"+code+"'", "destination")
    # not sold out
    print('https://www.seabourn.com/search/sbn_en_US/cruisesearch?&start=0&fq={!tag=destinationTag}destinationIds:(' + str(
            code) + ')&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&rows=10&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]')
    page = requests.get(
        'https://www.seabourn.com/search/sbn_en_US/cruisesearch?&start=0&fq={!tag=destinationTag}destinationIds:(' + str(
            code) + ')&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&rows=10&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]',
        headers=headers)
    cruise_results = page.json()
    total_results = int(cruise_results['results'])
    print("Found", total_results, "results")
    current_start = 0
    not_sold_out = 0
    sold_out = 0
    while current_start <= total_results:
        page = requests.get('https://www.seabourn.com/search/sbn_en_US/cruisesearch?&start=' + str(
            current_start) + '&fq={!tag=destinationTag}destinationIds:(' + str(
            code) + ')&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&rows=10&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]',
                            headers=headers)
        urls.append([code, page.json()['searchResults'], False])
        current_start += 10
    # Sold out
    page = requests.get(
        'https://www.seabourn.com/search/sbn_en_US/cruisesearch?&start=0&fq={!tag=destinationTag}destinationIds:(' + str(
            code) + ')&fq=(soldOut:(true)%20AND%20price_USD_anonymous:[1%20TO%20*])&rows=10&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]',
        headers=headers)
    cruise_results = page.json()
    total_results = int(cruise_results['results'])
    current_start = 0
    not_sold_out = 0
    sold_out = 0
    while current_start <= total_results:
        page = requests.get('https://www.seabourn.com/search/sbn_en_US/cruisesearch?&start=' + str(
            current_start) + '&fq={!tag=destinationTag}destinationIds:(' + str(
            code) + ')&fq=(soldOut:(true)%20AND%20price_USD_anonymous:[1%20TO%20*])&rows=10&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]',
                            headers=headers)
        urls.append([code, page.json()['searchResults'], True])
        current_start += 10


def convert_date(not_formatted):
    splitter = not_formatted.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def get_destination(param):
    if param == 'A':
        return ['Alaska', 'A']
    elif param == 'SN':
        return ['Antarctica & Patagonia', 'SN']
    elif param == 'I':
        return ['Arabia, Africa & India', 'I']
    elif param == 'O':
        return ['Asia', 'O']
    elif param == 'P':
        return ['Australia & South Pacific', 'P']
    elif param == 'N':
        return ['Canada & New England', 'NN']
    elif param == 'M':
        return ['Mediterranean', 'M']
    elif param == 'F':
        return ['Florida', 'F']
    elif param == 'C':
        return ['Caribbean', 'C']
    elif param == 'J':
        return ['Extended Explorations', 'J']
    elif param == 'EM':
        return ['Mediterranean', 'EM']
    elif param == 'X':
        return ['Holiday', 'X']
    elif param == 'EN':
        return ['Northern Europe', 'EN']
    elif param == 'L':
        return ['Pacific Coastal', 'L']
    elif param == 'T':
        return ['Panama Canal', 'T']
    elif param == 'S':
        return ['South America & Antarctica', 'S']
    elif param == 'ET':
        return ['Transatlantic', 'ET']
    elif param == 'H':
        return ['Dont Know', 'H']
    elif param == 'E':
        return ['Dont know', 'E']
    elif param == 'W':
        return ['World', 'W']
    elif param == 'Q':
        return ['Cuba', 'C']
    else:
        return ["MISSING.............................................................", param]


def get_vessel_id(name):
    if name == "Seabourn Encore":
        return "108"
    if name == "Seabourn Odyssey":
        return "580"
    if name == "Seabourn Ovation":
        return "926"
    if name == "Seabourn Quest":
        return "110"
    if name == "Seabourn Sojourn":
        return "719"


counter = 0
codes = []


def parse(ur):
    for result in ur[1]:
        brochure_name = result['title']
        vessel_name = result['shipName']
        cruise_line_name = "Seabourn"
        number_of_nights = (result['duration'])
        print("https://www.seabourn.com/api/v2/price/itinerary/" + result['itineraryId'])
        itinerary = requests.get("https://www.seabourn.com/api/v2/price/itinerary/" + result['itineraryId'],
                                 headers=price_headers).json()
        print("https://www.seabourn.com/api/v2/price/itinerary/" + result['itineraryId'])
        data = {}
        try:
            data = itinerary['data']
        except KeyError:
            print(itinerary)
        for sailing in data:
            interior_bucket_price = "N/A"
            balcony_bucket_price = "N/A"
            ocean_view_bucket_price = "N/A"
            suite_bucket_price = ""
            owner = ''
            spa = ''
            cruise_id = "8"
            destination = get_destination(ur[0])
            destination_name = destination[0]
            destination_code = destination[1]
            return_date = convert_date(sailing['arriveDate'])
            sail_date = convert_date(sailing['departDate'])
            vessel_id = get_vessel_id(vessel_name)
            if not ur[2]:
                for room in sailing['roomTypes']:
                    if room['id'].split('_')[1] == "OV":
                        if room['available']:
                            ocean_view_bucket_price = room['price'][0]['price']
                        else:
                            ocean_view_bucket_price = "N/A"
                    elif room['id'].split('_')[1] == "VS":
                        if room['available']:
                            balcony_bucket_price = room['price'][0]['price']
                        else:
                            balcony_bucket_price = "N/A"
                    elif room['id'].split('_')[1] == "PH":
                        if room['available']:
                            suite_bucket_price = room['price'][0]['price']
                        else:
                            suite_bucket_price = "N/A"
                    elif room['id'].split('_')[1] == "PH":
                        if room['available']:
                            owner = room['price'][0]['price']
                        else:
                            owner = "N/A"
                    elif room['id'].split('_')[1] == "PS":
                        if room['available']:
                            spa = room['price'][0]['price']
                        else:
                            spa = "N/A"
                if suite_bucket_price == "":
                    if spa == '':
                        if owner == '':
                            suite_bucket_price = "N/A"
                        else:
                            suite_bucket_price = owner
                    else:
                        suite_bucket_price = spa
                else:
                    pass
            else:
                interior_bucket_price = "N/A"
                ocean_view_bucket_price = "N/A"
                balcony_bucket_price = "N/A"
                suite_bucket_price = "N/A"
            temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name, "",
                    brochure_name, number_of_nights, sail_date, return_date, str(interior_bucket_price),
                    str(ocean_view_bucket_price), str(balcony_bucket_price), str(suite_bucket_price)]
            print(temp)
            if temp in to_write:
                pass
            else:
                to_write.append(temp)


pool.map(parse, urls)
pool.close()
pool.join()


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Seabourn.xlsx'
    if not os.path.exists(userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    workbook = xlsxwriter.Workbook(path_to_file)

    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})
    worksheet.set_column("A:A", 15)
    worksheet.set_column("B:B", 25)
    worksheet.set_column("C:C", 10)
    worksheet.set_column("D:D", 25)
    worksheet.set_column("E:E", 20)
    worksheet.set_column("F:F", 30)
    worksheet.set_column("G:G", 20)
    worksheet.set_column("H:H", 50)
    worksheet.set_column("I:I", 20)
    worksheet.set_column("J:J", 20)
    worksheet.set_column("K:K", 20)
    worksheet.set_column("L:L", 20)
    worksheet.set_column("M:M", 25)
    worksheet.set_column("N:N", 20)
    worksheet.set_column("O:O", 20)
    worksheet.write('A1', 'DestinationCode', bold)
    worksheet.write('B1', 'DestinationName', bold)
    worksheet.write('C1', 'VesselID', bold)
    worksheet.write('D1', 'VesselName', bold)
    worksheet.write('E1', 'CruiseID', bold)
    worksheet.write('F1', 'CruiseLineName', bold)
    worksheet.write('G1', 'ItineraryID', bold)
    worksheet.write('H1', 'BrochureName', bold)
    worksheet.write('I1', 'NumberOfNights', bold)
    worksheet.write('J1', 'SailDate', bold)
    worksheet.write('K1', 'ReturnDate', bold)
    worksheet.write('L1', 'InteriorBucketPrice', bold)
    worksheet.write('M1', 'OceanViewBucketPrice', bold)
    worksheet.write('N1', 'BalconyBucketPrice', bold)
    worksheet.write('O1', 'SuiteBucketPrice', bold)
    row_count = 1
    money_format = workbook.add_format({'bold': True})
    ordinary_number = workbook.add_format({"num_format": '#,##0'})
    date_format = workbook.add_format({'num_format': 'm d yyyy'})
    centered = workbook.add_format({'bold': True})
    money_format.set_align("center")
    money_format.set_bold(True)
    date_format.set_bold(True)
    centered.set_bold(True)
    ordinary_number.set_bold(True)
    ordinary_number.set_align("center")
    date_format.set_align("center")
    centered.set_align("center")
    for ship_entry in data_array:
        column_count = 0
        for en in ship_entry:
            if column_count == 0:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 1:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 2:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 3:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 4:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 5:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 6:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 7:
                worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 8:
                try:
                    worksheet.write_number(row_count, column_count, en, ordinary_number)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 9:
                try:
                    date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 10:
                try:
                    date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 11:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 12:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 13:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 14:
                try:
                    worksheet.write_number(row_count, column_count, int(en), money_format)
                except ValueError:
                    worksheet.write_string(row_count, column_count, en, centered)
            column_count += 1
        row_count += 1
    workbook.close()
    pass


write_file_to_excell(to_write)
print('Voyages:', len(to_write))
input("Press any key to continue...")
