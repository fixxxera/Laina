import datetime
import os

import requests
import xlsxwriter

session = requests.Session()
currency = "AUD"
data = {"CurrencyCode": currency, "PageSize": 200, "PageNumber": 1, "SortExpression": "FirstSailDate"}
page = session.post("https://www.carnival.com.au/DomainData/SailingSearch/Get/", data=data)
cruise_data = page.json()
data_array = []
special_list = []
tmp_legend_array = []
tmp_spirit_array = []


def check_if_correct(days_before_correct, ports):
    days = ports
    days_after_correct = days_before_correct
    for day in days:
        if day['PortName'] == 'X Intl Dateline':
            if days[0]['PortName'] == 'Sydney':
                days_after_correct = days_before_correct - 1
                return days_after_correct
            elif days[0]['PortName'] == 'Honolulu':
                days_after_correct = days_before_correct + 1
                return days_after_correct
    return days_after_correct


def convert_date(date):
    day = date[0]
    month = ''
    if date[1] == 'Jan':
        month = '1'
    elif date[1] == 'Feb':
        month = '2'
    elif date[1] == 'Mar':
        month = '3'
    elif date[1] == 'Apr':
        month = '4'
    elif date[1] == 'May':
        month = '5'
    elif date[1] == 'Jun':
        month = '6'
    elif date[1] == 'Jul':
        month = '7'
    elif date[1] == 'Aug':
        month = '8'
    elif date[1] == 'Sep':
        month = '9'
    elif date[1] == 'Oct':
        month = '10'
    elif date[1] == 'Nov':
        month = '11'
    elif date[1] == 'Dec':
        month = '12'
    year = date[2]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def get_date(corrected, date_text):
    result = []
    raw_date = (date_text.split())
    start_date = convert_date(raw_date)
    date_obj_start = datetime.datetime.strptime(start_date, "%m/%d/%Y")
    end_date = date_obj_start + datetime.timedelta(days=int(corrected))
    result.append(date_obj_start.strftime('%m/%d/%Y'))
    result.append(end_date.strftime('%m/%d/%Y'))
    return result


def match_by_meta(param):
    australia = ['Airlie Beach', 'Akaroa', 'Auckland', 'Bay Of Islands', 'Brisbane', 'Darwin', 'Fiordland Pk', 'Hobart',
                 'Melbourne', 'Mooloolaba', 'Moreton Island', 'Napier', 'Port Arthur', 'Port Douglas', 'Pt. Chalmers',
                 'Sydney', 'Tauranga', 'Wellington', 'Willis Island', 'Yorkeys Knob (Caims']
    exotics = ['Bali (Bebia)', 'Ho Chi Minh City (Ph', 'Ko Samui', 'Singapore']
    south_pacific_all = ['Bora Bora', 'Isle Of Pines', 'Lifou Isle', 'Mare', 'Moorea', 'Mystery Island', 'Noumea',
                         'Papeete', 'Port Denarau', 'Santo', 'Suva', 'Vila']
    ports_visited = param

    ports_list = []
    for i in range(len(ports_visited)):

        if i == 0:
            pass
        else:
            ports_list.append(ports_visited[i]['PortName'])
    result = []
    is_exotic = False
    is_pacific = False
    for element in exotics:
        if element in ports_list:
            is_exotic = True
    if not is_exotic:
        for element in south_pacific_all:
            if element in ports_list:
                is_pacific = True
    if not is_pacific:
        for element in australia:
            if element in ports_list:
                pass
    if is_exotic:
        result.append("Exotics")
        result.append("O")
        return result
    elif is_pacific:
        result.append("South Pacific -- All")
        result.append("I")
        return result
    else:
        result.append("Australia")
        result.append("P")
        return result


for row in cruise_data['Voyages']:
    is_special = False
    interior_price = row['FromIPrice']
    if interior_price != 'N/A':
        interior_price = interior_price.replace(" AUD", '').split('.')[0].replace(',', '')
    ocean_view = row['FromOPrice']
    if ocean_view != 'N/A':
        ocean_view = ocean_view.replace(" AUD", '').split('.')[0].replace(',', '')
    balcony = row['FromBPrice']
    if balcony != 'N/A':
        balcony = balcony.replace(" AUD", '').split('.')[0].replace(',', '')
    suite = row['FromSPrice']
    if suite != 'N/A':
        suite = suite.replace(" AUD", '').split('.')[0].replace(',', '')
    total_days = row['CruiseNights']
    corrected_days = check_if_correct(total_days, row['PortsVisited'])
    formatted_date = get_date(corrected_days, row['DateRangeText'])
    departure_date = formatted_date[0]
    arrival_date = formatted_date[1]
    ship_name = row['ShipName']
    title = row['VoyageTitle']
    destination = match_by_meta(row['PortsVisited'])
    dest = destination[0]
    destination_code = destination[1]
    ports_string = ''
    for port in row['PortsVisited']:
        print(port)
        ports_string += ", " + port['PortName']
    if ship_name == 'Legend':
        legend_array = [destination_code, dest, str(6), ship_name, str(2), 'Carnival Cruise Lines', '',
                        title,
                        total_days, departure_date, arrival_date, interior_price, ocean_view, balcony, suite, ports_string[1:]]
        tmp_legend_array.append(legend_array)
    else:
        spirit_array = [destination_code, dest, str(9), ship_name, str(2), 'Carnival Cruise Lines', '',
                        title,
                        total_days, departure_date, arrival_date, interior_price, ocean_view, balcony, suite, ports_string[1:]]
        tmp_spirit_array.append(spirit_array)
    for l in row['PortsVisited']:
        if 'IDL' in l['PortCode']:
            is_special = True
    if is_special:
        tmp = [destination_code, dest, str(6), ship_name, str(2), 'Carnival Cruise Lines', '', title, total_days,
               departure_date, arrival_date, interior_price, ocean_view, balcony, suite, ports_string[1:]]
        special_list.append(tmp)
data_array.append(tmp_legend_array)
data_array.append(tmp_spirit_array)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    # print(userhome)
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Carnival Australia.xlsx'
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
    worksheet.set_column("P:P", 21)
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
                        cell = ""
                        number = 0
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
                        cell = ""
                        number = 0
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
                        cell = ""
                        number = 0
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
                        cell = ""
                        number = 0
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


write_file_to_excell(data_array)


def write_special_to_excell(special_list):
    userhome = os.path.expanduser('~')
    # print(userhome)
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '- Carnival Australia special list.xlsx'
    if not os.path.exists(userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(
            now.month) + '-' + str(now.day)):
        os.makedirs(
            userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(now.day))
    print(path_to_file)
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
    worksheet.set_column("P:P", 20)
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
    row_count = 1
    rate = 1
    for item in special_list:
        column_count = 0
        for r in item:
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
                if '.' in r:
                    splitted = r.split('.')
                    strprice = splitted[0] + "" + splitted[1]
                else:
                    strprice = str(r)
                tmp = int(strprice) * rate
                if '.' in str(tmp):
                    splitted = str(tmp).split('.')
                    cell = int(splitted[0])
                else:
                    cell = int(tmp)
                worksheet.write_number(row_count, column_count, cell, money_format)
                column_count += 1
            elif column_count == 12:
                if '.' in r:
                    splitted = r.split('.')
                    strprice = splitted[0] + "" + splitted[1]
                else:
                    strprice = str(r)
                tmp = int(strprice) * rate
                if '.' in str(tmp):
                    splitted = str(tmp).split('.')
                    cell = int(splitted[0])
                else:
                    cell = int(tmp)
                worksheet.write_number(row_count, column_count, cell, money_format)
                column_count += 1
            elif column_count == 13:
                if '.' in r:
                    splitted = r.split('.')
                    strprice = splitted[0] + "" + splitted[1]
                else:
                    strprice = str(r)
                tmp = int(strprice) * rate
                if '.' in str(tmp):
                    splitted = str(tmp).split('.')
                    cell = int(splitted[0])
                else:
                    cell = int(tmp)
                worksheet.write_number(row_count, column_count, cell, money_format)
                column_count += 1
            elif column_count == 14:
                if '.' in r:
                    splitted = r.split('.')
                    strprice = splitted[0] + "" + splitted[1]
                else:
                    strprice = str(r)
                try:
                    tmp = int(strprice) * 1
                except ValueError:
                    pass
                if '.' in str(tmp):
                    splitted = str(tmp).split('.')
                    cell = int(splitted[0])
                else:
                    cell = int(tmp)
                worksheet.write_number(row_count, column_count, cell, money_format)
                column_count += 1
            elif column_count == 15:
                worksheet.write_string(row_count, column_count, str(r), centered)
                column_count += 1
        row_count += 1
    workbook.close()

write_special_to_excell(special_list)
input("Press any key to continue...")
