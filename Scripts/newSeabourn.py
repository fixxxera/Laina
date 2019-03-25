import datetime
import os
from time import sleep

import requests
from multiprocessing.dummy import Pool as ThreadPool

import xlsxwriter


def convert_date(not_formatted):
    splitter = not_formatted.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


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


headers_trips = {
    'Accept': 'application/json',
    'agencyId': 'Agency id',
    'Brand': 'sbn',
    'Content-Type': 'application/json',
    'country': 'US',
    'currencyCode': 'USD',
    'locale': 'en_US',
    'Referer': 'https://www.seabourn.com/en_US/find-a-cruise.html',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.109 Safari/537.36'
}
headers_prices = {
    'Accept': 'application/json',
    'agencyId': 'Agency id',
    'Brand': 'sbn',
    'Content-Type': 'application/json',
    'country': 'US',
    'currencyCode': 'USD',
    'locale': 'en_US',
    'Referer': 'https://www.seabourn.com/en_US/find-a-cruise.html',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.109 Safari/537.36'
}
pool = ThreadPool(10)
to_write = []
cruise_ids = []
prices = []
codes = ['A', 'SN', 'I', 'O', 'P', 'N', 'C', 'Q', 'J', 'X', 'EM', 'EN', 'L', 'T', 'S', 'ET', 'W']


def parse(code):
    sleep(2)
    code_result = requests.get(
        "https://www.seabourn.com/search/sbn_en_US/cruisesearch?&start=0&fq={!tag=destinationTag}destinationIds:(" + "\\" + code + ")&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&rows=500&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]")
    page = code_result.json()
    results = page['searchResults']
    for result in results:
        cruise_ids.append(result['cruiseId'].split("_")[0])
    for i in range(0, len(page['searchResults']), 100):
        all_ids = ""

        for index in range(i, i + 100):
            try:
                all_ids += cruise_ids[index]
                all_ids += ","
            except IndexError:
                break
        sleep(2)
        response = requests.get("https://www.seabourn.com/api/v2/price/cruise?cruiseIds=" + all_ids,
                                headers=headers_prices).json()['data']
        for sailing in response:
            prices.append(sailing)
    for result in results:
        ship_name = result['shipName']
        title = result['title']
        vessel_id = get_vessel_id(ship_name)
        duration = int(result['duration'])
        cid = result['cruiseId'].split('_')[0]
        start_date = convert_date(result['departureDate'].split('T')[0])
        end_date = convert_date(result['arrivalDate'].split('T')[0])
        destination = get_destination(code)
        destination_name = destination[0]
        destination_code = destination[1]
        interior = "N/A"
        oceanview = "N/A"
        veranda = "N/A"
        suite = "N/A"
        spa = "N/A"
        owners = "N/A"
        if not result['isSoldOut']:
            for price in prices:
                for room_type in price['roomTypes']:
                    if room_type['name'] == "Ocean View":
                        oceanview = room_type['lowestPrice']['price']
                        if oceanview == 0:
                            oceanview = "N/A"
                    elif room_type['name'] == "Vista Suite":
                        veranda = room_type['lowestPrice']['price']
                        if veranda == 0:
                            veranda = "N/A"
                    elif room_type['name'] == "Penthouse Suite":
                        suite = room_type['lowestPrice']['price']
                        if suite == 0:
                            suite = "N/A"
                    elif room_type['name'] == "Penthouse Spa Suite":
                        spa = room_type['lowestPrice']['price']
                        if spa == 0:
                            spa = "N/A"
            if suite == "N/A":
                if spa == "N/A":
                    if owners == "N/A":
                        suite = "N/A"
                    else:
                        suite = owners
                else:
                    suite = spa
            else:
                pass
        temp = [destination_code, destination_name, vessel_id, ship_name, cid, "Seabourn Cruises", "",
                title, duration, start_date, end_date, str(interior),
                str(oceanview), str(veranda), str(suite)]
        print(temp)
        if temp in to_write:
            pass
        else:
            to_write.append(temp)


pool.map(parse, codes)
pool.close()
pool.join()

print(len(prices))


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
