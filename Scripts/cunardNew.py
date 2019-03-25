import datetime
import os

import requests
from multiprocessing.dummy import Pool as ThreadPool

import xlsxwriter

headers = {
    "brand": "cunard",
    "country": "US",
    "currencycode": "USD",
    "locale": "en_US"
}
itineraries = []
to_write = []
destinations = ["Africa-and-Indian-Ocean", "Alaska", "Asia", "Australia-and-New-Zealand", "Canary-Islands", "Caribbean",
                "Mediterranean", "Northern-Europe", "Panama-Canal-and-Central-America", "South-America",
                "Transatlantic", "Usa-and-Canada", "Western-Europe", "World-Voyage", "Orient-Asia-Australia"]
pool = ThreadPool(5)
response = requests.get(
    "https://www.cunard.com/search/cunard_en_US/cruisesearch?&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&start=0&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&rows=6&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]&facet.field=flightPackage_USD_anonymous").json()
for destination in destinations:
    response = requests.get(
        "https://www.cunard.com/search/cunard_en_US/cruisesearch?&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&start=0&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&fq={!tag=destinationTag}destinationIds:(" + destination + ")&rows=" + str(
            response[
                'results']) + "&fq=departDate:[NOW/DAY%2B0DAY%20TO%20*]&facet.field=flightPackage_USD_anonymous").json()
    for result in response['searchResults']:
        itineraries.append([destination, result])


def convert_date(osd):
    splitter = osd.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (year, month, day)
    return final_date


def calculate_days(sail_date_param, number_of_nights_param):
    date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
    try:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))
    except ValueError:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param.split("-")[1]))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def parse(itinerary):
    path = itinerary[1]['itineraryPath']
    cid = itinerary[1]['itineraryId']
    title = itinerary[1]['title']
    ship_name = itinerary[1]['shipName']
    sold_out = itinerary[1]['isSoldOut']
    if not sold_out:
        prices = requests.get("https://www.cunard.com/api/v2/price/cruise/" + cid + "?", headers=headers).json()['data']
        cid = prices['cruiseCode']
        start_date = convert_date(prices['departDate'])
        end_date = convert_date(prices['arriveDate'])
        duration = prices['duration']
        for roomtype in prices['roomTypes']:
            if roomtype['id'] in "B":
                balcony = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "S":
                mini_suite = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "I":
                interior = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "O":
                outside = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "A":
                club = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "Q":
                queens = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "BB":
                balcony = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "BO":
                outside = roomtype['lowestPrice']['price']
            elif roomtype['id'] in "BI":
                interior = roomtype['lowestPrice']['price']
        temp = [cid, itinerary[0], "", ship_name, cid, "Cunard Cruises", "",
                title, duration, start_date, end_date, str(interior),
                str(outside), str(balcony), str(mini_suite)]
        print(temp)
        if temp in to_write:
            pass
        else:
            to_write.append(temp)


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


pool.map(parse, itineraries)
pool.close()
pool.join()
