import datetime
import os
from json import JSONDecodeError

import requests
from multiprocessing.dummy import Pool as ThreadPool

import xlsxwriter

session = requests.Session()
destinations = ["AUS", "KHM", "COK", "TLS", "FJI", "IDN", "MYS", "NCL", "NZL", "NFK", "PNG", "SGP", "SLB", "THA", "TON",
                "VUT", "WSM"]
headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive"
}
body = {
    "searchParameters": {
        "p": [],
        "c": [],
        "r": False,
        "d": [],
        "s": [],
        "ms": [],
        "adv": [],
        "sort": "dpa",
        "page": 1,
        "rt": [],
        "CorrelationId": 1492857025880},
    "renderingParameters": {
        "AdditionalPromoCodes": [],
        "CruiseItinerary": [],
        "DefaultSortOption": {"Code": "dpa"},
        "DeparturePort": [],
        "Duration": [],
        "ExcludeVoyage": [],
        "FareTypes": [],
        "KeepPageState": True,
        "LastMinuteDeal": False,
        "MaxNumberOfResults": 0,
        "MidWeekDeparture": False,
        "NumberOfAdults": 0,
        "NumberOfChildren": 0,
        "NumberOfInfants": 0,
        "OnSale": False,
        "PromoCode": [],
        "RoomType": [],
        "SchoolHoliday": False,
        "Ship": [],
        "VisitingCountry": [],
        "VisitingPort": [],
        "Voyage": [],
        "WeekendDeparture": False}}
to_write = []
pool = ThreadPool(5)
sailings = []
codes = set()
itineraries_count = 0


def calculate_days(date, duration):
    dateobj = datetime.datetime.strptime(date, "%m/%d/%Y")
    calculated = dateobj + datetime.timedelta(days=int(duration))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def get_vessel_id(ves_name):
    if ves_name == "Pacific Aria":
        return "1"
    elif ves_name == "Pacific Dawn":
        return "1"
    elif ves_name == "Pacific Eden":
        return "1"
    elif ves_name == "Pacific Explorer":
        return "1"
    elif ves_name == "Pacific Jewel":
        return "1"
    elif ves_name == "Pacific Pearl":
        return "1"
    else:
        return '1'


for dest in destinations:
    current_page = 1
    body['searchParameters']['c'] = [dest]
    body['searchParameters']['page'] = current_page
    main_url = "https://www.pocruises.com.au/sc_ignore/b2c/cruiseresults/searchresultsV2"
    page = session.post(main_url, json=body, headers=headers)
    cruise_data = page.json()
    total_pages = cruise_data['MetaData']['PageCount']
    while current_page <= total_pages:
        body['searchParameters']['page'] = current_page
        main_url = "https://www.pocruises.com.au/sc_ignore/b2c/cruiseresults/searchresultsV2"
        page = session.post(main_url, json=body, headers=headers)
        cruise_data = page.json()
        for item in cruise_data['Items']:
            itineraries_count += 1
            for voyage in item['Voyages']:
                voyage_code = voyage['VoyageCode']
                if voyage_code in codes:
                    continue
                number_of_nights = voyage['NumberOfNights']
                raw_date = datetime.datetime.fromtimestamp(
                    int(voyage['DepartureDate'].replace('/Date(', '').replace(')/', '')) / 1000)
                sail_date = str(raw_date.month) + "/" + str(raw_date.day) + "/" + str(raw_date.year)
                return_date = calculate_days(sail_date, str(number_of_nights))
                brochure_name = voyage['Title']
                vessel_name = voyage['Ship']
                vessel_id = get_vessel_id(vessel_name)
                json = {"voyage": voyage_code, "adults": 2, "childrenDateOfBirths": []}
                main_url = "https://www.pocruises.com.au/sc_ignore/b2c/BookingJson/GetFares"
                page = session.post(main_url, json=json, headers=headers)
                cruise_data = page.json()
                interior = ""
                oceanview = ""
                balcony = ""
                mini_suite = ""
                suite = ""
                cruise_id = ''
                cruise_line_name = "P&O AU"
                package_id = ''
                for current_price in cruise_data['Result']:
                    if current_price['Title'] == 'Interior':
                        if current_price['SoldOut']:
                            interior = "N/A"
                        else:
                            for fare in current_price['FareTypes']:
                                if fare['DefaultFare']:
                                    interior = str(fare['LeadInFare']['PerPersonPrice'])
                                if '.' in interior:
                                    interior = interior.split('.')[0]
                    elif current_price['Title'] == 'Oceanview':
                        if current_price['SoldOut']:
                            oceanview = "N/A"
                        else:
                            for fare in current_price['FareTypes']:
                                if fare['DefaultFare']:
                                    oceanview = str(fare['LeadInFare']['PerPersonPrice'])
                                if '.' in oceanview:
                                    oceanview = oceanview.split('.')[0]
                    elif current_price['Title'] == 'Balcony':
                        if current_price['SoldOut']:
                            balcony = "N/A"
                        else:
                            for fare in current_price['FareTypes']:
                                if fare['DefaultFare']:
                                    balcony = str(fare['LeadInFare']['PerPersonPrice'])
                                if '.' in balcony:
                                    balcony = balcony.split('.')[0]
                    elif current_price['Title'] == 'Mini Suite':
                        if current_price['SoldOut']:
                            mini_suite = "N/A"
                        else:
                            for fare in current_price['FareTypes']:
                                if fare['DefaultFare']:
                                    mini_suite = str(fare['LeadInFare']['PerPersonPrice'])
                                if '.' in mini_suite:
                                    mini_suite = mini_suite.split('.')[0]
                    elif current_price['Title'] == 'Suite':
                        if current_price['SoldOut']:
                            suite = "N/A"
                        else:
                            for fare in current_price['FareTypes']:
                                if fare['DefaultFare']:
                                    suite = str(fare['LeadInFare']['PerPersonPrice'])
                                if '.' in suite:
                                    suite = suite.split('.')[0]
                    if mini_suite == '':
                        mini_suite = suite
                    if mini_suite == '':
                        mini_suite = "N/A"
                codes.add(voyage_code)
                temp = [dest, dest, vessel_id, vessel_name, cruise_id, cruise_line_name,
                        package_id,
                        brochure_name, number_of_nights, sail_date, return_date,
                        interior, oceanview, balcony, mini_suite]
                temp2 = [temp]
                to_write.append(temp2)
                print("Sailing:", temp)
                # Get siblings and parse
                cruise_data = session.get(
                    'https://www.pocruises.com.au/sc_ignore/b2c/cruiseprofile/Siblings?numberOfColumns=100&voyageCode=' + voyage_code,
                    headers=headers)
                try:
                    for column in cruise_data.json()['Data']:
                        for sibling in column:
                            if sibling['VoyageCode'] in codes:
                                continue
                            else:
                                voyage_code = sibling['VoyageCode']
                                if sibling['VoyageCode'] in codes:
                                    continue
                                else:
                                    codes.add(sibling['VoyageCode'])
                                if sibling['Sailed']:
                                    codes.add(sibling['VoyageCode'])
                                    continue
                                raw_date = datetime.datetime.fromtimestamp(
                                    int(sibling['DepartureDate'].replace('/Date(', '').replace(')/', '')) / 1000)
                                sail_date = str(raw_date.month) + "/" + str(raw_date.day) + "/" + str(raw_date.year)
                                return_date = calculate_days(sail_date, str(number_of_nights))
                                vessel_name = sibling['ShipName']
                                vessel_id = get_vessel_id(vessel_name)
                                json = {"voyage": sibling['VoyageCode'], "adults": 2, "childrenDateOfBirths": []}
                                main_url = "https://www.pocruises.com.au/sc_ignore/b2c/BookingJson/GetFares"
                                page = session.post(main_url, json=json, headers=headers)
                                cruise_data = page.json()
                                interior = ""
                                oceanview = ""
                                balcony = ""
                                mini_suite = ""
                                suite = ""
                                cruise_id = ''
                                cruise_line_name = "P&O UK"
                                package_id = ''
                                try:
                                    for current_price in cruise_data['Result']:
                                        if current_price['Title'] == 'Interior':
                                            if current_price['SoldOut']:
                                                interior = "N/A"
                                            else:
                                                for fare in current_price['FareTypes']:
                                                    if fare['DefaultFare']:
                                                        interior = str(fare['LeadInFare']['PerPersonPrice'])
                                                    if '.' in interior:
                                                        interior = interior.split('.')[0]
                                        elif current_price['Title'] == 'Oceanview':
                                            if current_price['SoldOut']:
                                                oceanview = "N/A"
                                            else:
                                                for fare in current_price['FareTypes']:
                                                    if fare['DefaultFare']:
                                                        oceanview = str(fare['LeadInFare']['PerPersonPrice'])
                                                    if '.' in oceanview:
                                                        oceanview = oceanview.split('.')[0]
                                        elif current_price['Title'] == 'Balcony':
                                            if current_price['SoldOut']:
                                                balcony = "N/A"
                                            else:
                                                for fare in current_price['FareTypes']:
                                                    if fare['DefaultFare']:
                                                        balcony = str(fare['LeadInFare']['PerPersonPrice'])
                                                    if '.' in balcony:
                                                        balcony = balcony.split('.')[0]
                                        elif current_price['Title'] == 'Mini Suite':
                                            if current_price['SoldOut']:
                                                mini_suite = "N/A"
                                            else:
                                                for fare in current_price['FareTypes']:
                                                    if fare['DefaultFare']:
                                                        mini_suite = str(fare['LeadInFare']['PerPersonPrice'])
                                                    if '.' in mini_suite:
                                                        mini_suite = mini_suite.split('.')[0]
                                        elif current_price['Title'] == 'Suite':
                                            if current_price['SoldOut']:
                                                suite = "N/A"
                                            else:
                                                for fare in current_price['FareTypes']:
                                                    if fare['DefaultFare']:
                                                        suite = str(fare['LeadInFare']['PerPersonPrice'])
                                                    if '.' in suite:
                                                        suite = suite.split('.')[0]
                                        if mini_suite == '':
                                            mini_suite = suite
                                        if mini_suite == '':
                                            mini_suite = "N/A"
                                    codes.add(voyage_code)
                                    temp = [dest, dest, vessel_id, vessel_name, cruise_id, cruise_line_name,
                                            package_id,
                                            brochure_name, number_of_nights, sail_date, return_date,
                                            interior, oceanview, balcony, mini_suite]
                                    temp2 = [temp]
                                    to_write.append(temp2)
                                    print("    Sibling:", temp)
                                except TypeError:
                                    print(main_url)
                except JSONDecodeError:
                    continue
        current_page += 1


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- P&O AU.xlsx'
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
                except ValueError:
                    worksheet.write_string(row_count, column_count, r, centered)
                    column_count += 1
            row_count += 1
    workbook.close()
    pass
write_file_to_excell(to_write)
input("Press any key to continue...")
