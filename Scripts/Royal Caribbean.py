import datetime
import os

import requests
from multiprocessing.dummy import Pool as ThreadPool

import xlsxwriter

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive"
}
to_crawl = []
to_parse = []
errored = []
to_write = []
mini_list = []
init_response = requests.get(
    'https://secure.royalcaribbean.com/ajax/cruises/service/lookup?initialLoad=true&icid=hpnvbr_pagehe_fnd_hm_awaren_938&loadLinks=true',
    headers=headers).json()
total_pages = init_response['listResultsModule']['totalPages']
to_crawl.append(init_response['listResultsModule']['resultData']['pageResults'])

for i in range(0, total_pages):
    if i == 0:
        to_crawl.append(
            'https://secure.royalcaribbean.com/ajax/cruises/service/lookup?initialLoad=true&icid=hpnvbr_pagehe_fnd_hm_awaren_938&loadLinks=true')
    else:
        to_crawl.append(
            'https://secure.royalcaribbean.com/ajax/cruises/service/lookup?initialLoad=false&currentPage=' + str(
                i) + '&icid=hpnvbr_pagehe_fnd_hm_awaren_938&loadLinks=true')


def crawl(page):
    try:
        response = requests.get(page, headers=headers)
        try:
            to_parse.append(response.json()['listResultsModule']['resultData']['pageResults'])
        except KeyError:
            pass
    except requests.exceptions.InvalidSchema:
        errored.append(page)


def get_destination(param):
    if param == 'ALCAN':
        return ['Alaska', 'A']
    elif param == 'FAR.E':
        return ['Exotics', 'O']
    elif param == 'AUSTL':
        return ['AU/NZ', 'P']
    elif param == 'BAHAM':
        return ['Bahamas', 'BH']
    elif param == 'BERMU':
        return ['Bermuda', 'BM']
    elif param == 'ATLCO':
        return ['Can/New En', 'NN']
    elif param == 'CARIB':
        return ['Carib', 'C']
    elif param == 'CUBAN':
        return ['Cuba', 'C']
    elif param == 'EUROP':
        return ['Europe', 'E']
    elif param == 'HAWAI':
        return ['Hawaii', 'H']
    elif param == 'PACIF':
        return ['Pacific', 'I']
    elif param == 'ISLAN':
        return ['Repositioning', 'R']
    elif param == 'SOPAC':
        return ['South Pacific', 'I']
    elif param == 'T.ATL':
        return ['Transatlantic', 'E']
    elif param == 'TPACI':
        return ['Transpacific', 'I']
    elif param == 'DUBAI':
        return ['DUBAI', 'DU']
    elif param == 'T.PAN':
        return ['T.PAN', 'T.PAN']


def calculate_days(sail_date_param, number_of_nights_param):
    date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
    try:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))
    except ValueError:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param.split("-")[1]))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def convert_date(osd):
    splitter = osd.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


pool = ThreadPool(5)
pool.map(crawl, to_crawl)
pool.close()
pool.join()


def get_vessel_id(ves_name):
    if ves_name == "Anthem of the Seas":
        return "859"
    elif ves_name == "Ovation of the Seas":
        return "931"
    elif ves_name == "Quantum of the Seas":
        return "860"
    elif ves_name == "Allure of the Seas":
        return "717"
    elif ves_name == "Harmony of the Seas":
        return "941"
    elif ves_name == "Oasis of the Seas":
        return "691"
    elif ves_name == "Freedom of the Seas":
        return "502"
    elif ves_name == "Independence of the Seas":
        return "581"
    elif ves_name == "Liberty of the Seas":
        return "561"
    elif ves_name == "Adventure of the Seas":
        return "378"
    elif ves_name == "Explorer of the Seas":
        return "217"
    elif ves_name == "Mariner of the Seas":
        return "417"
    elif ves_name == "Navigator of the Seas":
        return "408"
    elif ves_name == "Voyager of the Seas":
        return "237"
    elif ves_name == "Brilliance of the Seas":
        return "399"
    elif ves_name == "Jewel of the Seas":
        return "432"
    elif ves_name == "Radiance of the Seas":
        return "225"
    elif ves_name == "Serenade of the Seas":
        return "416"
    elif ves_name == "Enchantment of the Seas":
        return "216"
    elif ves_name == "Grandeur of the Seas":
        return "218"
    elif ves_name == "Legend of the Seas":
        return "219"
    elif ves_name == "Rhapsody of the Seas":
        return "228"
    elif ves_name == "Vision of the Seas":
        return "236"
    elif ves_name == "Majesty of the Seas":
        return "220"
    elif ves_name == "Empress of the Seas":
        return "999"
    else:
        return ""


def split_carib(ports, dn, dc):
    wc = ['Costa Maya, Mexico', 'Cozumel, Mexico', 'Falmouth, Jamaica', 'George Town, Grand Cayman',
          'Ocho Rios, Jamaica']

    ec = ['Basseterre, St. Kitts', 'Bridgetown, Barbados', 'Castries, St. Lucia', 'Charlotte Amalie, St. Thomas',
          'Fort De France', 'Kingstown, St. Vincent', 'Philipsburg, St. Maarten', 'Ponce, Puerto Rico',
          'Punta Cana, Dominican Rep', 'Roseau, Dominica', 'San Juan, Puerto Rico', 'St. Croix, U.S.V.I.',
          "St. George's, Grenada", "St. John's, Antigua", 'Tortola, B.V.I']
    ports_list = []
    for i in range(len(ports)):

        if i == 0:
            pass
        else:
            ports_list.append(ports[i])

    for element in wc:
        for p in ports_list:
            if p in element or element in p:
                return ['West Carib', 'C']

    for element in ec:
        for p in ports_list:
            if p in element or element in p:
                return ['East Carib', 'C']

    return [dn, dc]


def split_repo(ports, dn, dc):
    wc = ['Costa Maya, Mexico', 'Cozumel, Mexico', 'Falmouth, Jamaica', 'George Town, Grand Cayman',
          'Ocho Rios, Jamaica']

    ec = ['Basseterre, St. Kitts', 'Bridgetown, Barbados', 'Castries, St. Lucia', 'Charlotte Amalie, St. Thomas',
          'Fort De France', 'Kingstown, St. Vincent', 'Philipsburg, St. Maarten', 'Ponce, Puerto Rico',
          'Punta Cana, Dominican Rep', 'Roseau, Dominica', 'San Juan, Puerto Rico', 'St. Croix, U.S.V.I.',
          "St. George's, Grenada", "St. John's, Antigua", 'Tortola, B.V.I']
    nn = ['Halifax', 'Charlottetown']
    ports_list = []
    for i in range(len(ports)):

        if i == 0:
            pass
        else:
            ports_list.append(ports[i])

    for element in wc:
        for p in ports_list:
            if p in element or element in p:
                return ['West Carib', 'C']

    for element in ec:
        for p in ports_list:
            if p in element or element in p:
                return ['East Carib', 'C']

    for element in nn:
        for p in ports_list:
            if p in element or element in p:
                return ['Can/New En', 'NN']
    return [dn, dc]


def split_europe(ports, dn, dc):
    baltic = ['Petropavlovsk, Russia', 'Bergen, Norway', 'Flam, Norway', 'Geiranger, Norway', 'Alesund, Norway',
              'Stavanger, Norway', 'Skjolden, Norway', 'Stockholm, Sweden', 'Helsinki, Finland',
              'St. Petersburg, Russia', 'Tallinn, Estonia', 'Riga, Latvia', 'Warnemunde, Germany',
              'Copenhagen, Denmark', 'Kristiansand, Norway', 'Skagen, Denmark', 'Fredericia, Denmark',
              'Rostock (Berlin), Germany', 'Nynashamn, Sweden', 'Oslo, Norway', 'Amsterdam, Netherlands',
              'Reykjavik, Iceland',
              'Zeebrugge (Brussels), Belgium', 'Southampton, England']
    eastern_med = ['Athens (Piraeus), Greece', 'Katakolon, Greece', 'Dubrovnik, Croatia', 'Mykonos, Greece',
                   'Rhodes, Greece', 'Chania (Souda),Crete, Greece', 'Koper, Slovenia', 'Split, Croatia',
                   'Santorini, Greece', 'Zadar, Croatia', 'Corfu, Greece', 'Kotor, Montenegro']
    west_med = ['Catania,Sicily,Italy', 'Ajaccio, Corsica', 'Alicante, Spain', 'Barcelona, Spain', 'Bilbao, Spain',
                'Cadiz, Spain', 'Cannes, France', 'Cartagena, Spain', 'Florence / Pisa (Livorno),Italy',
                'Fuerteventura, Canary', 'Funchal (Madeira), Portugal', 'Genoa, Italy', 'Gibraltar, United Kingdom',
                'Ibiza, Spain', 'La Coruna, Spain', 'La Spezia, Italy', 'Lanzarote, Canary Islands',
                'Las Palmas, Gran Canaria', 'Lisbon, Portugal', 'Malaga, Spain', 'Marseille, France',
                'Messina (Sicily), Italy', 'Montecarlo, Monaco', 'Naples, Italy', 'Nice (Villefranche)',
                'Palma De Mallorca, Spain', 'Ponta Delgada, Azores', 'Portofino, Italy', 'Provence (Toulon), France',
                'Ravenna, Italy', 'Sete, France', 'St. Peter Port, Channel Isl', 'Tenerife, Canary Islands',
                'Valencia, Spain', 'Valletta, Malta', 'Venice, Italy', 'Vigo, Spain']

    ports_visited = ports

    ports_list = []
    for i in range(len(ports_visited)):

        if i == 0:
            pass
        else:
            ports_list.append(ports_visited[i])
    for element in baltic:
        for p in ports_list:
            if p in element or element in p:
                return ['Baltic', 'E']
            elif ports_visited[0] in element or element in ports_visited[0]:
                return ['Baltic', 'E']

    for element in eastern_med:
        for p in ports_list:
            if p in element or element in p:
                return ['Eastern Med', 'E']

    for element in west_med:
        for p in ports_list:
            if p in element or element in p:
                return ['Western Med', 'E']

    return [dn, dc]


for entry in to_parse:
    print(entry)
    for itinerary in entry:
        cruise_line_name = "Royal Caribbean"
        destination = get_destination(itinerary['destinationCode'])
        destination_name = destination[0]
        destination_code = destination[1]
        title = itinerary['displayName']
        duration = itinerary['totalNights']
        vessel_name = itinerary['ship']['name']
        vessel_id = get_vessel_id(vessel_name)
        inside = 'N/A'
        balcony = 'N/A'
        oceanview = 'N/A'
        suite = 'N/A'
        for sailing in itinerary['sailings']:
            ports = []
            for port in sailing['cruisePorts']:
                ports.append(port['name'])
            cruise_id = "14"
            for price in sailing['priceRanges']:
                if price['category'] == 'DELUXE':
                    suite = str(int(price['amount']))
                    if '.' in suite:
                        suite = suite.split('.')[0]
                elif price['category'] == 'BALCONY':
                    balcony = str(int(price['amount']))
                    if '.' in balcony:
                        balcony = balcony.split('.')[0]
                elif price['category'] == 'OUTSIDE':
                    oceanview = str(int(price['amount']))
                    if '.' in oceanview:
                        suite = oceanview.split('.')[0]
                elif price['category'] == 'INTERIOR':
                    inside = str(int(price['amount']))
                    if '.' in inside:
                        inside = inside.split('.')[0]
            sail_date = convert_date(sailing['startDate'])
            return_date = calculate_days(sail_date, duration)
            if 'Carib' in destination_name:
                dest = split_carib(ports, destination_name, destination_code)
                destination_code = dest[1]
                destination_name = dest[0]
            if destination_name == 'Carib':
                if 'Western Caribbean' in title:
                    destination_name = 'West Carib'
            if 'Repositioning' in destination_name:
                dest = split_repo(ports, destination_name, destination_code)
                destination_code = dest[1]
                destination_name = dest[0]
            if 'Europe' in destination_name:
                dest = split_europe(ports, destination_name, destination_code)
                destination_code = dest[1]
                destination_name = dest[0]
            final_ports = ''
            for p in ports:
                final_ports += str(p + ', ')
            if destination_code == 'O':
                if 'Hong Kong' in ports[0] or 'Shenzhen' in ports[0] or 'Tianjin' in ports[0] or 'Shanghai' in ports[0]:
                    if "(" in ports[0]:
                        destination_name = ports[0].split()[0]
                    else:
                        destination_name = ports[0]
            final_ports = final_ports.strip()[:-1]

            temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
                    "", title, duration, sail_date, return_date, inside,
                    oceanview, balcony, suite, final_ports]
            print(temp)
            mini_list.append(temp)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Royal Caribbean.xlsx'
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
    worksheet.set_column("P:P", 30)
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
    worksheet.write('P1', 'Ports', bold)
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
            if column_count == 15:
                worksheet.write_string(row_count, column_count, en, centered)
            column_count += 1
        row_count += 1
    workbook.close()
    pass


write_file_to_excell(mini_list)
