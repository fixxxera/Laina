import datetime
import os

import xlsxwriter
from requests import get
from multiprocessing.dummy import Pool as ThreadPool

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Host": "www.hollandamerica.com",
    "currencycode": "USD",
    "country": "US",
    "Refer": "https://www.hollandamerica.com/en_US/find-a-cruise.html"
}
headers_secondary = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Host": "www.hollandamerica.com",
    "currencycode": "USD",
    "country": "US",
    "Refer": "https://www.hollandamerica.com/en_US/find-a-cruise.html",
    "brand": "hal",
    "locale": "en_US"
}
pool = ThreadPool(5)


def convert_date(not_formatted):
    splitter = not_formatted.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def get_vessel_id(name):
    if name == "Amsterdam":
        return "108"
    if name == "Eurodam":
        return "580"
    if name == "Koningsdam":
        return "926"
    if name == "Maasdam":
        return "110"
    if name == "Nieuw Amsterdam":
        return "719"
    if name == "Noordam":
        return "496"
    if name == "Oosterdam":
        return "410"
    if name == "Prinsendam":
        return "407"
    if name == "Rotterdam":
        return "113"
    if name == "Veendam":
        return "118"
    if name == "Volendam":
        return "119"
    if name == "Westerdam":
        return "434"
    if name == "Zaandam":
        return "121"
    if name == "Zuiderdam":
        return "409"


def get_destination(param):
    if param == 'CA':
        return ['Caribbean Cuba', 'C']
    elif param == 'CE':
        return ['Caribbean Eastern', 'C']
    elif param == 'CF':
        return ['Caribbean Panama Canal', 'C']
    elif param == 'CS':
        return ['Caribbean Southern', 'C']
    elif param == 'CT':
        return ['Caribbean Tropical', 'C']
    elif param == 'CW':
        return ['Caribbean Western', 'C']
    elif param == '4D1':
        return ['Alaska Denali', '4D1']
    elif param == '4D2':
        return ['Denali 2days', '4D2']
    elif param == '4D3':
        return ['Denali 3days', '4D3']
    elif param == '4Y2':
        return ['2days Denali&Yukon', '4Y2']
    elif param == '4Y3':
        return ['3days Denali&Yukon', '4Y3']
    elif param == 'EM':
        return ['Europe Mediterranean', 'EM']
    elif param == 'EN':
        return ['Europe Northern', 'EN']
    elif param == 'ET':
        return ['Europe Transatlantic', 'ET']
    elif param == 'GB1':
        return ['GlacierBay', 'GB1']
    elif param == 'WA':
        return ['Grand Asia & Australia', 'WA']
    elif param == 'WS':
        return ['Grand South America & Antarctica', 'WS']
    elif param == 'WW':
        return ['Grand World Voyages', 'WW']
    elif param == 'HUB':
        return ['Hubbard Glacier', 'HUB']
    elif param == 'SN':
        return ['South America & Antarctica', 'SN']
    elif param == 'SS':
        return ['South America', 'SS']
    elif param == 'TAC':
        return ['Tracy Arm Sawyer Gracier', 'TAC']


count = get(
    'https://www.hollandamerica.com/search/hal_en_US/cruisesearch?&start=0&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&rows=10&fq=departDate:[NOW/DAY%2B1DAY%20TO%20*]',
    headers=headers)
print(count.json())
print("Total itineraries: ", count.json()['results'])
count_saver = int(count.json()['results'])
all_itineraries = []
for page in range(0, int(count.json()['results']), 100):

    if count_saver < 100:
        print("Downloading itineraries ", page, " - ", page + count_saver)
        current_page = get(
            'https://www.hollandamerica.com/search/hal_en_US/cruisesearch?&start=' + str(
                count_saver) + '&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&rows=100&fq=departDate:[NOW/DAY%2B1DAY%20TO%20*]',
            headers=headers).json()
    else:
        print("Downloading itineraries ", page, " - ", page + 100)
        current_page = get(
            'https://www.hollandamerica.com/search/hal_en_US/cruisesearch?&start=' + str(
                page) + '&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&rows=100&fq=departDate:[NOW/DAY%2B1DAY%20TO%20*]',
            headers=headers).json()
    count_saver -= 100
    all_itineraries.append(current_page)
to_write = []


def parse(item):
    for result in item['searchResults']:
        path = result['itineraryPath']
        title = result['title']
        if "ms " in result['shipName']:
            vessel_name = result['shipName'].split()[1]
        else:
            vessel_name = result['shipName']
        duration = result['duration']
        code = result['itineraryId']
        cruise_id = "8"
        try:
            destination = get_destination(result['regions'][0])
            destination_name = destination[0]
            destination_code = destination[1]
        except IndexError:
            destination_name = "No Info"
            destination_code = "No Info"

        vessel_id = get_vessel_id(vessel_name)
        inside = ''
        oceanview = ''
        verandah = ''
        suite = ''
        neptune = ''
        lanai = ''
        signature = ''
        sail_date = ''
        return_date = ''
        sold_out = result['isSoldOut']
        if sold_out:
            inside = 'N/A'
            oceanview = 'N/A'
            verandah = 'N/A'
            suite = 'N/A'
            neptune = 'N/A'
            lanai = 'N/A'
            signature = 'N/A'
            for entry in \
                    get('https://www.hollandamerica.com/api/v2/price/itinerary/' + code,
                        headers=headers_secondary).json()[
                        'data']:
                sail_date = convert_date(entry['departDate'])
                return_date = convert_date(entry['arriveDate'])
            if vessel_name == "Amsterdam" or vessel_name == 'Maasdam' or vessel_name == 'Prinsendam' or vessel_name == 'Rotterdam' or vessel_name == 'Veendam' or vessel_name == 'Volendam' or vessel_name == 'Zaandam':
                if suite != '':
                    verandah = suite
                else:
                    verandah = 'N/A'
                if signature != '':
                    suite = signature
                else:
                    if neptune != '':
                        suite = neptune
                    else:
                        suite = "N/A"
            elif vessel_name == "Eurodam" or vessel_name == 'Nieuw Amsterdam' or vessel_name == 'Noordam' or vessel_name == 'Oosterdam' or vessel_name == 'Prinsendam' or vessel_name == 'Westerdam' or vessel_name == 'Zuiderdam':
                if signature == '':
                    suite = neptune
                else:
                    suite = signature
            elif vessel_name == 'Koningsdam':
                if verandah == '':
                    if suite == '':
                        verandah = 'N/A'
                    else:
                        verandah = suite
                if suite == "":
                    if signature == '':
                        if neptune == '':
                            suite = "N/A"
                        else:
                            suite = neptune
                    else:
                        suite = signature
                else:
                    pass
            temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, "Holland America", "",
                    title, duration, sail_date, return_date, str(inside),
                    str(oceanview), str(verandah), str(suite)]
            print(temp)
            if temp in to_write:
                pass
            else:
                to_write.append(temp)
        else:
            sailings = []
            url = 'https://www.hollandamerica.com/api/v2/price/itinerary/' + code
            try:
                print(url)
                sailings = get(url, headers=headers_secondary).json()['data']
            except KeyError:
                print(url)
            for entry in sailings:
                sail_date = convert_date(entry['departDate'])
                return_date = convert_date(entry['arriveDate'])
                for room in entry['roomTypes']:
                    if room['name'] == 'Inside':
                        if room['available']:
                            inside = room['price'][0]['price']
                        else:
                            inside = 'N/A'
                    elif room['name'] == 'Ocean View':
                        if room['available']:
                            oceanview = room['price'][0]['price']
                        else:
                            oceanview = 'N/A'
                    elif room['name'] == 'Lanai':
                        if room['available']:
                            lanai = room['price'][0]['price']
                        else:
                            lanai = 'N/A'
                    elif room['name'] == 'Vista Suite':
                        if room['available']:
                            suite = room['price'][0]['price']
                        else:
                            suite = 'N/A'
                    elif room['name'] == 'Neptune Suite':
                        if room['available']:
                            neptune = room['price'][0]['price']
                        else:
                            neptune = 'N/A'
                    elif room['name'] == 'Verandah':
                        if room['available']:
                            verandah = room['price'][0]['price']
                        else:
                            verandah = 'N/A'
                    elif room['name'] == 'Signature Suite':
                        if room['available']:
                            signature = room['price'][0]['price']
                        else:
                            signature = 'N/A'
                    else:
                        print(room['name'])
                if vessel_name == "Amsterdam" or vessel_name == 'Maasdam' or vessel_name == 'Prinsendam' or vessel_name == 'Rotterdam' or vessel_name == 'Veendam' or vessel_name == 'Volendam' or vessel_name == 'Zaandam':
                    if suite != '':
                        verandah = suite
                    else:
                        verandah = 'N/A'
                    if signature != '':
                        suite = signature
                    else:
                        if neptune != '':
                            suite = neptune
                        else:
                            suite = "N/A"
                elif vessel_name == "Eurodam" or vessel_name == 'Nieuw Amsterdam' or vessel_name == 'Noordam' or vessel_name == 'Oosterdam' or vessel_name == 'Prinsendam' or vessel_name == 'Westerdam' or vessel_name == 'Zuiderdam':
                    if signature == '':
                        suite = neptune
                    else:
                        suite = signature
                elif vessel_name == 'Koningsdam':
                    if verandah == '':
                        if suite == '':
                            verandah = 'N/A'
                        else:
                            verandah = suite
                    if suite == "":
                        if signature == '':
                            if neptune == '':
                                suite = "N/A"
                            else:
                                suite = neptune
                        else:
                            suite = signature
                    else:
                        pass
                temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, "Holland America", "",
                        title, duration, sail_date, return_date, str(inside),
                        str(oceanview), str(verandah), str(suite)]
                print(temp)
                if temp in to_write:
                    pass
                else:
                    to_write.append(temp)


pool.map(parse, all_itineraries)
pool.close()
pool.join()


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Holland America.xlsx'
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
                try:
                    worksheet.write_string(row_count, column_count, en, centered)
                except TypeError:
                    worksheet.write_string(row_count, column_count, " ", centered)
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
