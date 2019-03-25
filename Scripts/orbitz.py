import datetime
import json
import math
import os
from multiprocessing.dummy import Pool as ThreadPool

import requests
import xlrd as xlrd
import xlsxwriter

session = requests.session()
pool = ThreadPool(2)
page_requests = []
codes = []
all_cruises = []

headers = {
    'Host': 'www.orbitz.com',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:63.0) Gecko/20100101 Firefox/63.0',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
    'Cache-Control': 'max-age=0',
    'TE': 'Trailers',
    'Refer': 'https://www.orbitz.com/Cruise-Search?destination=&earliest-departure-date=&_xpid=11905%7C1&adultCount=2',
    'X-Requested-With': 'XMLHttpRequest'
}

url = 'https://www.orbitz.com/cruise/search/api/sailings?adultCount=2&childCount=0&offset=1&page=1'
response = session.get(url, headers=headers).json()
print("Total results:", response['searchSailingResponseType']['total'])
pages = int(math.ceil(response['searchSailingResponseType']['total'] / 25))
current_page = 1
while current_page <= pages:
    page_requests.append(
        'https://www.orbitz.com/cruise/search/api/sailings?offset=' + str(current_page) + '&adults=2&page=' + str(
            current_page))
    current_page += 1
print(len(page_requests))
workbook = xlrd.open_workbook('CruisePrices.xlsx', on_demand=True)
worksheet = workbook.sheet_by_index(0)
first_row = []
for col in range(worksheet.ncols):
    first_row.append(worksheet.cell_value(0, col))
data = []
for row in range(0, worksheet.nrows):
    elm = {}
    for col in range(worksheet.ncols):
        elm[first_row[col]] = worksheet.cell_value(row, col)
    data.append(elm)


def convert_date(param):
    splitter = param.split('-')
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    if month == 'Jan':
        month = '1'
    elif month == 'Feb':
        month = '2'
    elif month == 'Mar':
        month = '3'
    elif month == 'Apr':
        month = '4'
    elif month == 'May':
        month = '5'
    elif month == 'Jun':
        month = '6'
    elif month == 'Jul':
        month = '7'
    elif month == 'Aug':
        month = '8'
    elif month == 'Sep':
        month = '9'
    elif month == 'Oct':
        month = '10'
    elif month == 'Nov':
        month = '11'
    elif month == 'Dec':
        month = '12'
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def calculate_days(date, duration):
    dateobj = datetime.datetime.strptime(date, "%m/%d/%Y")
    calculated = dateobj + datetime.timedelta(days=int(duration))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def get_destination(param):
    if param == 'Europe':
        return 'E'
    elif param == 'Asia':
        return 'O'
    elif param == 'Transatlantic':
        return 'Transatlantic'
    elif param == 'World':
        return 'World'
    elif param == 'Middle East':
        return 'Middle East'
    elif param == 'Africa':
        return 'Africa'
    elif param == 'South Pacific':
        return 'I'
    elif param == 'Alaska':
        return 'A'
    elif param == 'Arctic / Antarctic':
        return 'Arctic / Antarctic'
    elif param == 'South America':
        return 'S'
    elif param == 'Africa':
        return 'Africa'
    elif param == 'Getaway at Sea':
        return 'Getaway at Sea'
    elif param == 'Transpacific':
        return 'Transpacific'
    elif param == 'Australia / New Zealand':
        return 'P'
    elif param == 'Caribbean':
        return 'C'
    elif param == 'Mexico':
        return 'MX'
    elif param == 'Bahamas':
        return 'BH'
    elif param == 'Bermuda':
        return 'BM'
    elif param == 'Pacific Coastal':
        return 'Pacific Coastal'
    elif param == 'Canada / New England':
        return 'NN'
    elif param == 'Hawaii':
        return 'H'
    elif param == 'Panama Canal':
        return 'T'
    else:
        return param


def get_from_vessel_name(vessel_name):
    if vessel_name == 'Adventure of the Seas':
        return '109'
    elif vessel_name == 'Allure of the Seas':
        return "510"
    elif vessel_name == 'Amsterdam':
        return "58"
    elif vessel_name == 'Anthem of the Seas':
        return "663"
    elif vessel_name == 'Azamara Journey':
        return "408"
    elif vessel_name == 'Azamara Quest':
        return "437"
    elif vessel_name == 'Brilliance of the Seas':
        return "111"
    elif vessel_name == 'Caribbean Princess':
        return "89"
    elif vessel_name == 'Carnival Breeze':
        return "542"
    elif vessel_name == 'Carnival Conquest':
        return "3"
    elif vessel_name == 'Carnival Dream':
        return "491"
    elif vessel_name == 'Carnival Ecstasy':
        return "14"
    elif vessel_name == 'Carnival Elation':
        return "15"
    elif vessel_name == 'Carnival Fantasy':
        return "16"
    elif vessel_name == 'Carnival Fascination':
        return "17"
    elif vessel_name == 'Carnival Freedom':
        return "182"
    elif vessel_name == 'Carnival Glory':
        return "5"
    elif vessel_name == 'Carnival Horizon':
        return "28"
    elif vessel_name == 'Carnival Imagination':
        return "19"
    elif vessel_name == 'Carnival Inspiration':
        return "20"
    elif vessel_name == 'Carnival Legend':
        return "6"
    elif vessel_name == 'Carnival Liberty':
        return "142"
    elif vessel_name == 'Carnival Magic':
        return "518"
    elif vessel_name == 'Carnival Miracle':
        return "7"
    elif vessel_name == 'Carnival Paradise':
        return "22"
    elif vessel_name == 'Carnival Pride':
        return "8"
    elif vessel_name == 'Carnival Sensation':
        return "23"
    elif vessel_name == 'Carnival Splendor':
        return "449"
    elif vessel_name == 'Carnival Sunshine':
        return "4"
    elif vessel_name == 'Carnival Triumph':
        return "10"
    elif vessel_name == 'Carnival Valor':
        return "11"
    elif vessel_name == 'Carnival Victory':
        return "12"
    elif vessel_name == 'Carnival Vista':
        return "697"
    elif vessel_name == 'Celebrity Constellation':
        return "26"
    elif vessel_name == 'Celebrity Eclipse':
        return "711"
    elif vessel_name == 'Celebrity Edge':
        return "DontKnow"
    elif vessel_name == 'Celebrity Equinox':
        return "443"
    elif vessel_name == 'Celebrity Infinity':
        return "29"
    elif vessel_name == 'Celebrity Millennium':
        return "31"
    elif vessel_name == 'Celebrity Reflection':
        return "540"
    elif vessel_name == 'Celebrity Silhouette':
        return "525"
    elif vessel_name == 'Celebrity Solstice':
        return "444"
    elif vessel_name == 'Celebrity Summit':
        return "32"
    elif vessel_name == 'Celebrity Xpedition':
        return "33"
    elif vessel_name == 'Celebrity Xperience':
        return "732"
    elif vessel_name == 'Celebrity Xploration':
        return "733"
    elif vessel_name == 'Coral Princess':
        return "90"
    elif vessel_name == 'Costa Deliziosa':
        return "509"
    elif vessel_name == 'Costa Diadema':
        return "649"
    elif vessel_name == 'Costa Fascinosa':
        return "565"
    elif vessel_name == 'Costa Favolosa':
        return "520"
    elif vessel_name == 'Costa Luminosa':
        return "501"
    elif vessel_name == 'Costa Magica':
        return "41"
    elif vessel_name == 'Costa Mediterranea':
        return "42"
    elif vessel_name == 'Costa neoClassica':
        return "38"
    elif vessel_name == 'Costa neoRiviera':
        return "688"
    elif vessel_name == 'Costa Pacifica':
        return "502"
    elif vessel_name == 'Crown Princess':
        return "165"
    elif vessel_name == 'Diamond Princess':
        return "92"
    elif vessel_name == 'Disney Dream':
        return "517"
    elif vessel_name == 'Disney Fantasy':
        return "532"
    elif vessel_name == 'Disney Magic':
        return "55"
    elif vessel_name == 'Disney Wonder':
        return "56"
    elif vessel_name == 'Emerald Princess':
        return "186"
    elif vessel_name == 'Empress of the Seas':
        return "112"
    elif vessel_name == 'Enchantment of the Seas':
        return "113"
    elif vessel_name == 'Eurodam':
        return "441"
    elif vessel_name == 'Explorer of the Seas':
        return "114"
    elif vessel_name == 'Freedom of the Seas':
        return "158"
    elif vessel_name == 'Golden Princess':
        return "93"
    elif vessel_name == 'Grand Princess':
        return "94"
    elif vessel_name == 'Grandeur of the Seas':
        return "115"
    elif vessel_name == 'Harmony of the Seas':
        return "700"
    elif vessel_name == 'Independence of the Seas':
        return "445"
    elif vessel_name == 'Island Princess':
        return "95"
    elif vessel_name == 'Jewel of the Seas':
        return "116"
    elif vessel_name == 'Koningsdam':
        return "692"
    elif vessel_name == 'Liberty of the Seas':
        return "184"
    elif vessel_name == 'Maasdam':
        return "59"
    elif vessel_name == 'Majestic Princess':
        return "730"
    elif vessel_name == 'Majesty of the Seas':
        return "118"
    elif vessel_name == 'Mariner of the Seas':
        return "119"
    elif vessel_name == 'MSC Armonia':
        return "159"
    elif vessel_name == 'MSC Bellissima':
        return "DontKnow"
    elif vessel_name == 'MSC Divina':
        return "547"
    elif vessel_name == 'MSC Fantasia':
        return "493"
    elif vessel_name == 'MSC Magnifica':
        return "505"
    elif vessel_name == 'MSC Meraviglia':
        return "718"
    elif vessel_name == 'MSC Musica':
        return "167"
    elif vessel_name == 'MSC Orchestra':
        return "410"
    elif vessel_name == 'MSC Poesia':
        return "469"
    elif vessel_name == 'MSC Preziosa':
        return "615"
    elif vessel_name == 'MSC Seaside':
        return "717"
    elif vessel_name == 'MSC Seaview':
        return "734"
    elif vessel_name == 'MSC Sinfonia':
        return "164"
    elif vessel_name == 'MSC Splendida':
        return "506"
    elif vessel_name == 'Navigator of the Seas':
        return "121"
    elif vessel_name == 'Nieuw Amsterdam':
        return "514"
    elif vessel_name == 'Nieuw Statendam':
        return "65"
    elif vessel_name == 'Noordam':
        return "60"
    elif vessel_name == 'Norwegian Bliss':
        return "DontKnow"
    elif vessel_name == 'Norwegian Breakaway':
        return "548"
    elif vessel_name == 'Norwegian Dawn':
        return "73"
    elif vessel_name == 'Norwegian Epic':
        return "513"
    elif vessel_name == 'Norwegian Escape':
        return "662"
    elif vessel_name == 'Norwegian Gem':
        return "185"
    elif vessel_name == 'Norwegian Getaway':
        return "606"
    elif vessel_name == 'Norwegian Jade':
        return "450"
    elif vessel_name == 'Norwegian Jewel':
        return "157"
    elif vessel_name == 'Norwegian Joy':
        return "DontKnow"
    elif vessel_name == 'Norwegian Pearl':
        return "183"
    elif vessel_name == 'Norwegian Sky':
        return "77"
    elif vessel_name == 'Norwegian Spirit':
        return "78"
    elif vessel_name == 'Norwegian Star':
        return "79"
    elif vessel_name == 'Norwegian Sun':
        return "80"
    elif vessel_name == 'Oasis of the Seas':
        return "498"
    elif vessel_name == 'Oosterdam':
        return "61"
    elif vessel_name == 'Ovation of the Seas':
        return "701"
    elif vessel_name == 'Pacific Princess':
        return "96"
    elif vessel_name == 'Pride of America':
        return "144"
    elif vessel_name == 'Prinsendam':
        return "62"
    elif vessel_name == 'Quantum of the Seas':
        return "644"
    elif vessel_name == 'Queen Elizabeth':
        return "512"
    elif vessel_name == 'Queen Mary 2':
        return "53"
    elif vessel_name == 'Queen Victoria':
        return "188"
    elif vessel_name == 'Radiance of the Seas':
        return "122"
    elif vessel_name == 'Regal Princess':
        return "97"
    elif vessel_name == 'Rhapsody of the Seas':
        return "123"
    elif vessel_name == 'Rotterdam':
        return "63"
    elif vessel_name == 'Royal Princess':
        return "98"
    elif vessel_name == 'Ruby Princess':
        return "460"
    elif vessel_name == 'Sapphire Princess':
        return "99"
    elif vessel_name == 'Sea Princess':
        return "143"
    elif vessel_name == 'Serenade of the Seas':
        return "124"
    elif vessel_name == 'Star Princess':
        return "100"
    elif vessel_name == 'Sun Princess':
        return "101"
    elif vessel_name == 'Symphony of the Seas':
        return "DontKnow"
    elif vessel_name == 'Veendam':
        return "66"
    elif vessel_name == 'Vision of the Seas':
        return "127"
    elif vessel_name == 'Volendam':
        return "67"
    elif vessel_name == 'Voyager of the Seas':
        return "128"
    elif vessel_name == 'Westerdam':
        return "68"
    elif vessel_name == 'Zaandam':
        return "69"
    elif vessel_name == 'Zuiderdam':
        return "70"
    else:

        return " "


def xldate_to_datetime(xldate):
    try:
        tempDate = datetime.datetime(1900, 1, 1)
        deltaDays = datetime.timedelta(days=int(xldate))
        secs = (int((xldate % 1) * 86400) - 60)
        detlaSeconds = datetime.timedelta(seconds=secs)
        TheTime = (tempDate + deltaDays + detlaSeconds)
        return TheTime.strftime("%m/%d/%Y")
    except ValueError:
        old_value = xldate.split('/')
        new_value = (
                datetime.date(int(old_value[2]), int(old_value[0]), int(old_value[1])) - datetime.date(1899, 12,
                                                                                                       30)).days
        tempDate = datetime.datetime(1900, 1, 1)
        deltaDays = datetime.timedelta(days=int(new_value))
        secs = (int((new_value % 1) * 86400) - 60)
        detlaSeconds = datetime.timedelta(seconds=secs)
        TheTime = (tempDate + deltaDays + detlaSeconds)
        return TheTime.strftime("%m/%d/%Y")


def parse(page):
    try:
        resp = requests.get(page, headers=headers).json()
        itineraries = resp['searchSailingResponseType']['sailings']
        # sponsored = resp['sponsoredListing']['itineraries']
    except KeyError:
        print('Retrying', page)
        try:
            resp = requests.get(page, headers=headers).json()
            itineraries = resp['searchSailingResponseType']['sailings']
            # sponsored = resp['sponsoredListing']['itineraries']
        except KeyError:
            print('Skipping', page)
            return
    except json.decoder.JSONDecodeError:
        print('Retrying', page)
        try:
            resp = requests.get(page, headers=headers).json()
            itineraries = resp['searchSailingResponseType']['sailings']
            # sponsored = resp['sponsoredListing']['itineraries']
        except KeyError:
            print('Skipping', page)
            return
    for i in itineraries:
        cruise_line_name = i['cruiseLine']['name']
        number_of_days = i['length']
        vessel_name = i['ship']['name']
        vessel_id = get_from_vessel_name(vessel_name)
        destination = i['itinerary']['destination']
        dest = get_destination(destination['destination'])
        sub_dest = destination['subDestination']

        if i['sailingCode'] in codes:
            continue
        else:
            codes.append(i['sailingCode'])
        ports = []
        for location in i['itinerary']['locations']:
            if location['location']['type'] == "PORT":
                try:
                    ports.append(location['location']['name'])
                except KeyError:
                    ports.append(location['location']['countryName'] + "waters")
        ports_for_write = ports
        sail_date = convert_date(i['departureDate'])
        return_date = calculate_days(sail_date, number_of_days)
        prices = []
        try:
            prices = i['cabinClasses']
        except KeyError:
            print(i)
        interior = ''
        oceanview = ''
        balcony = ''
        suite = ''
        for p in prices:
            if p['code'] == 1:
                if p['leadInPrice'] == '' or p['leadInPrice'] is None:
                    interior = 'N/A'
                else:
                    interior = str(int(str(p['leadInPrice']).strip().replace(' ', '').split('.')[0]) / 2).split('.')[0]
            elif p['code'] == 2:
                if p['leadInPrice'] == '' or p['leadInPrice'] is None:
                    oceanview = 'N/A'
                else:
                    oceanview = str(int(str(p['leadInPrice']).strip().replace(' ', '').split('.')[0]) / 2).split('.')[0]
            elif p['code'] == 3:
                if p['leadInPrice'] == '' or p['leadInPrice'] is None:
                    balcony = 'N/A'
                else:
                    balcony = str(int(str(p['leadInPrice']).strip().replace(' ', '').split('.')[0]) / 2).split('.')[0]
            elif p['code'] == 4:
                if p['leadInPrice'] == '' or p['leadInPrice'] is None:
                    suite = 'N/A'
                else:
                    suite = str(int(str(p['leadInPrice']).strip().replace(' ', '').split('.')[0]) / 2).split('.')[0]
            if interior == '':
                interior = 'N/A'
            if oceanview == '':
                oceanview = 'N/A'
            if balcony == '':
                balcony = 'N/A'
            if suite == '':
                suite = 'N/A'
        for d in data:
            if d['VesselName'] == vessel_name and xldate_to_datetime(
                    d['SailDate']) == sail_date and xldate_to_datetime(d['ReturnDate']) == return_date:
                if d['DestinationCode'] == dest and d['DestinationName'] == sub_dest:
                    pass
                else:
                    dest = d['DestinationCode']
                    sub_dest = d['DestinationName']
                    break
            else:
                pass
        if sub_dest == 'BALTICS':
            sub_dest = "Baltics"
        if sub_dest == 'WEST MED':
            sub_dest = 'WMED'
        if sub_dest == 'EAST MED':
            sub_dest = 'EMED'
        if dest == "CUBA":
            dest = "C"
        temp = [dest, sub_dest, vessel_id, vessel_name, '', cruise_line_name, '',
                '', number_of_days, sail_date, return_date,
                interior, oceanview, balcony, suite, ports_for_write]
        temp2 = [temp]
        print(temp)
        all_cruises.append(temp2)
    # for i in sponsored:
    #     cruise_line_name = i['cruiseLine']['name']
    #     number_of_days = i['length']
    #     vessel_name = i['ship']['name']
    #     vessel_id = get_from_vessel_name(vessel_name)
    #     destination = i['itinerary']['destination']
    #     dest = get_destination(destination['destination'])
    #     sub_dest = destination['subDestination']
    #     sailings = i['sailings']
    #
    #     for s in sailings:
    #         if s['sailingCode'] in codes:
    #             continue
    #         else:
    #             codes.append(s['sailingCode'])
    #         portlist = s['locations']
    #         ports = []
    #         for l in portlist:
    #             if l['location']['type'] == 'PORT':
    #                 try:
    #                     ports.append(l['location']['name'])
    #                 except KeyError:
    #                     ports.append(l['location']['countryName'] + "waters")
    #         ports_for_write = ports
    #         sail_date = convert_date(s['departureDate'])
    #         return_date = calculate_days(sail_date, number_of_days)
    #         prices = s['leadInPriceBycode']
    #         interior = ''
    #         oceanview = ''
    #         balcony = ''
    #         suite = ''
    #         for p in prices:
    #             if p['code'] == '1':
    #                 if p['price'] == '' or p['price'] is None:
    #                     interior = 'N/A'
    #                 else:
    #                     interior = int(str(p['price']).strip().replace(' ', '').split('.')[0]) / 2
    #             elif p['code'] == '2':
    #                 if p['price'] == '' or p['price'] is None:
    #                     oceanview = 'N/A'
    #                 else:
    #                     oceanview = int(str(p['price']).strip().replace(' ', '').split('.')[0]) / 2
    #             elif p['code'] == '3':
    #                 if p['price'] == '' or p['price'] is None:
    #                     balcony = 'N/A'
    #                 else:
    #                     balcony = int(str(p['price']).strip().replace(' ', '').split('.')[0]) / 2
    #             elif p['code'] == '4':
    #                 if p['price'] == '' or p['price'] is None:
    #                     suite = 'N/A'
    #                 else:
    #                     suite = int(str(p['price']).strip().replace(' ', '').split('.')[0]) / 2
    #             if interior == '':
    #                 interior = 'N/A'
    #             if oceanview == '':
    #                 oceanview = 'N/A'
    #             if balcony == '':
    #                 balcony = 'N/A'
    #             if suite == '':
    #                 suite = 'N/A'
    #         for d in data:
    #             if d['VesselName'] == vessel_name and xldate_to_datetime(
    #                     d['SailDate']) == sail_date and xldate_to_datetime(d['ReturnDate']) == return_date:
    #                 if d['DestinationCode'] == dest and d['DestinationName'] == sub_dest:
    #                     pass
    #                 else:
    #                     dest = d['DestinationCode']
    #                     sub_dest = d['DestinationName']
    #                     break
    #             else:
    #                 pass
    #         if sub_dest == 'BALTICS':
    #             sub_dest = "Baltics"
    #         if sub_dest == 'WEST MED':
    #             sub_dest = 'WMED'
    #         if sub_dest == 'EAST MED':
    #             sub_dest = 'EMED'
    #         if dest == "CUBA":
    #             dest = "C"
    #         temp = [dest, sub_dest, vessel_id, vessel_name, '', cruise_line_name, '',
    #                 '', number_of_days, sail_date, return_date,
    #                 interior, oceanview, balcony, suite, ports_for_write]
    #         temp2 = [temp]
    #         print(temp)
    #         all_cruises.append(temp2)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Orbitz.xlsx'
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
    worksheet.set_column("P:P", 100)
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
    worksheet.write('P1', 'PortList', bold)
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
    for ship in data_array:
        for l in ship:
            column_count = 0
            for r in l:
                try:
                    if column_count == 0:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 1:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 2:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 3:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 4:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 5:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 6:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 7:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                    elif column_count == 8:
                        worksheet.write_number(row_count, column_count, int(r), ordinary_number)
                        column_count += 1
                    elif column_count == 9:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 10:
                        date_time = datetime.datetime.strptime(str(r), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, centered)
                        column_count += 1
                    elif column_count == 11:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 12:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 13:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 14:
                        tmp = str(r)
                        if "." in tmp:
                            number = round(float(tmp))
                        else:
                            number = int(tmp)
                        if number == 0:
                            cell = "N/A"
                            worksheet.write(row_count, column_count, cell, centered)
                        else:
                            worksheet.write_number(row_count, column_count, number, money_format)
                        column_count += 1
                    elif column_count == 15:
                        worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                except ValueError:
                    worksheet.write_string(row_count, column_count, r, centered)
                    column_count += 1

            row_count += 1
    workbook.close()
    pass


pool.map(parse, page_requests)
pool.close()
pool.join()
write_file_to_excell(all_cruises)
input("Press any key to continue...")
