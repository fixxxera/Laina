import datetime
import os
from multiprocessing.dummy import Pool as ThreadPool

import requests
import xlsxwriter
from bs4 import BeautifulSoup

session = requests.session()
pool = ThreadPool(5)
resp = session.get("https://www.azamaraclubcruises.com/int/json/find-a-voyage/search")
results = []
for voyage in resp.json()['voyages']:
    results.append(voyage)


def convert_date(day, month, year):
    if month == 'January' or month == 'Jan':
        month = '1'
    elif month == 'February' or month == 'Feb':
        month = '2'
    elif month == 'March' or month == 'Mar':
        month = '3'
    elif month == 'April' or month == 'Apr':
        month = '4'
    elif month == 'May' or month == 'May':
        month = '5'
    elif month == 'June' or month == 'Jun':
        month = '6'
    elif month == 'July' or month == 'Jul':
        month = '7'
    elif month == 'August' or month == 'Aug':
        month = '8'
    elif month == 'September' or month == 'Sep':
        month = '9'
    elif month == 'October' or month == 'Oct':
        month = '10'
    elif month == 'November' or month == 'Nov':
        month = '11'
    elif month == 'December' or month == 'Dec':
        month = '12'
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def calculate_days(date, duration):
    dateobj = datetime.datetime.strptime(date, "%m/%d/%Y")
    calculated = dateobj + datetime.timedelta(days=int(duration))
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


all_cruises = []


def get_destination(text):
    if text == 'Asia':
        return ['Exotics', 'O']
    elif text == 'Australia & New Zealand':
        return ['Australia', 'P']
    elif text == 'Mediterranean':
        return ['Mediterranean', 'E']
    elif text == 'Northern & Western Europe':
        return ['Europe', 'E']
    elif text == 'Panama Canal, Central & North America':
        return ['South America', 'S']
    elif text == 'Alaska, Panama Canal & North America':
        return ['South America', 'S']
    elif text == 'Cuba & Caribbean':
        return ['Carib', 'C']


def split_carib(ports, dn, dc):
    cu = ['Santiago de Cuba', 'Cienfuegos', 'Havana']
    wc = ['Costa Maya, Mexico', 'Cozumel, Mexico', 'Falmouth, Jamaica', 'George Town, Grand Cayman',
          'Ocho Rios, Jamaica']

    ec = ['Basseterre, St. Kitts', 'Bridgetown, Barbados', 'Castries, St. Lucia', 'Charlotte Amalie, St. Thomas',
          'Fort De France', 'Kingstown, St. Vincent', 'Philipsburg, St. Maarten', 'Ponce, Puerto Rico',
          'Punta Cana, Dominican Rep', 'Roseau, Dominica', 'San Juan, Puerto Rico', 'St. Croix, U.S.V.I.',
          "St. George's, Grenada", "St. John's, Antigua", 'Tortola, B.V.I', 'Tortola']
    ports_list = []
    for i in range(len(ports)):

        if i == 0:
            pass
        else:
            ports_list.append(ports[i])
    for element in cu:
        for p in ports_list:
            if p in element or element in p:
                return ['Cuba', 'C']
    for element in wc:
        for p in ports_list:
            if p in element or element in p:
                return ['West Carib', 'C']

    for element in ec:
        for p in ports_list:
            if p in element or element in p:
                return ['East Carib', 'C']

    return [dn, dc]


def split_europe(ports, dn, dc):
    baltic = ['Petropavlovsk', 'Bergen', 'Flam', 'Geiranger', 'Alesund',
              'Stavanger', 'Skjolden', 'Stockholm', 'Helsinki, Finland',
              'St. Petersburg', 'Tallinn', 'Riga', 'Warnemunde',
              'Copenhagen', 'Kristiansand', 'Skagen', 'Fredericia',
              'Rostock (Berlin)', 'Nynashamn', 'Oslo', 'Amsterdam',
              'Reykjavik',
              'Zeebrugge (Brussels), Belgium', 'Southampton']
    eastern_med = ['Athens (Piraeus)', 'Limassol, Cyprus', 'Katakolon', 'Dubrovnik', 'Mykonos',
                   'Rhodes', 'Chania (Souda)', 'Crete', 'Koper, Slovenia', 'Split',
                   'Santorini', 'Zadar', 'Corfu', 'Kotor']
    west_med = ['Catania,Sicily', 'Ajaccio, Corsica', 'Alicante', 'Barcelona', 'Bilbao',
                'Cadiz', 'Cannes', 'Cartagena', 'Florence / Pisa (Livorno)',
                'Fuerteventura, Canary', 'Funchal (Madeira)', 'Genoa', 'Gibraltar',
                'Ibiza', 'La Coruna', 'La Spezia', 'Lanzarote, Canary Islands',
                'Las Palmas, Gran Canaria', 'Lisbon', 'Malaga', 'Marseille',
                'Messina (Sicily)', 'Montecarlo', 'Naples', 'Nice (Villefranche)',
                'Palma De Mallorca', 'Ponta Delgada, Azores', 'Portofino', 'Provence (Toulon)',
                'Ravenna', 'Sete', 'St. Peter Port, Channel Isl', 'Tenerife, Canary Islands',
                'Valencia', 'Valletta', 'Venice', 'Vigo']
    europe = ['Rome (Civitavecchia)', 'Le Havre (Paris)', 'Akureyri',
              'Belfast, Northern Ireland', 'Cherbourg', 'Cork (Cobh)', 'Dover',
              'Dublin', 'Edinburgh', 'Greenock (Glasgow)', 'Inverness/Loch Ness',
              'Lerwick/Shetland', 'Liverpool',
              'Waterford (Dunmore E.)']

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
    for element in europe:
        for p in ports_list:
            if p in element or element in p:
                return ['Europe', 'E']

    return ['Europe', 'E']


def split_australia(ports):
    p = ['Adelaide, Australia', 'Airlie Beach, Qld, Australia', 'Alotau',
         'Ben Boyd National Park (Scenic Cruising Port)', 'Broome', 'Burnie', 'Busselton', 'Cairns, Australia',
         'Cooktown', 'Eden', 'Esperance, Australia', 'Exmouth', 'Fiordland National Park (Scenic Cruising)',
         'Fraser Island', 'Perth (Fremantle), Australia', 'Geraldton', 'Gladstone', 'Hamilton Island',
         'Hobart, Tasmania', 'Kangaroo Island', 'Kimberley Coast (Scenic Cruising Port)', 'Mooloolaba - Sunshine Coast',
         'Mornington Peninsula', 'Napier, New Zealand', 'Picton, New Zealand', 'Port Lincoln', 'Portland, Maine',
         'Stewart Island', 'Sydney Harbour Mooring (Scenic Cruising Port)', 'Townsville',
         'White Island (Scenic Cruising Port)', 'Wilsons Promontory (Scenic Cruising Port)']

    o = ['Benoa', 'Ko Chang', 'Komodo Island', 'Krabi', 'Bangkok/Laem Chabang, Thailand', 'Langkawi', 'Lombok',
         'Makassar', 'Phuket, Thailand', 'Probolinggo', 'Sabang (Palau Weh)', 'Sihanoukville']

    ip = ['Apia, Samoa', 'Conflict Islands', "Dili - Timor L'Este", 'Dravuni Island', 'Gizo Island', 'Honiara',
          'Kawanasausau Strait & Milne Bay (Scenic Cruising Port)', 'Kiriwina Island', 'Kitava',
          'Lifou, Loyalty Island', 'Madang',
          'Mutiny on the Bounty (Scenic Cruising Port)', "Nuku 'alofa, Tonga", 'Rabaul', 'Santo',
          "Vavau (Neiafu), Tonga", 'Vitu Islands (Scenic Cruising Port)', 'Wewak', 'Lahaina (Maui), Hawaii']
    ports_list = []
    for i in range(len(ports)):

        if i == 0:
            pass
        else:
            ports_list.append(ports[i])
    result = []
    is_exotic = False
    is_pacific = False
    for element in o:
        if element in ports_list:
            is_exotic = True
    if not is_exotic:
        for element in ip:
            if element in ports_list:
                is_pacific = True
    if not is_pacific:
        for element in p:
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
        result.append("Australia/New Zealand")
        result.append("P")
        return result


def check_carib(ports, dn, dc):
    baltic = ['Petropavlovsk', 'Bergen', 'Flam', 'Geiranger', 'Alesund',
              'Stavanger', 'Skjolden', 'Stockholm', 'Helsinki, Finland',
              'St. Petersburg', 'Tallinn', 'Riga', 'Warnemunde',
              'Copenhagen', 'Kristiansand', 'Skagen', 'Fredericia',
              'Rostock (Berlin)', 'Nynashamn', 'Oslo', 'Amsterdam',
              'Reykjavik',
              'Zeebrugge (Brussels), Belgium', 'Southampton']
    eastern_med = ['Athens (Piraeus)', 'Limassol, Cyprus', 'Katakolon', 'Dubrovnik', 'Mykonos',
                   'Rhodes', 'Chania (Souda)', 'Crete', 'Koper, Slovenia', 'Split',
                   'Santorini', 'Zadar', 'Corfu', 'Kotor']
    west_med = ['Catania,Sicily', 'Ajaccio, Corsica', 'Alicante', 'Barcelona', 'Bilbao',
                'Cadiz', 'Cannes', 'Cartagena', 'Florence / Pisa (Livorno)',
                'Fuerteventura, Canary', 'Funchal (Madeira)', 'Genoa', 'Gibraltar',
                'Ibiza', 'La Coruna', 'La Spezia', 'Lanzarote, Canary Islands',
                'Las Palmas, Gran Canaria', 'Lisbon', 'Malaga', 'Marseille',
                'Messina (Sicily)', 'Montecarlo', 'Naples', 'Nice (Villefranche)',
                'Palma De Mallorca', 'Ponta Delgada, Azores', 'Portofino', 'Provence (Toulon)',
                'Ravenna', 'Sete', 'St. Peter Port, Channel Isl', 'Tenerife, Canary Islands',
                'Valencia', 'Valletta', 'Venice', 'Vigo']
    europe = ['Rome (Civitavecchia)', 'Le Havre (Paris)', 'Akureyri',
              'Belfast, Northern Ireland', 'Cherbourg', 'Cork (Cobh)', 'Dover',
              'Dublin', 'Edinburgh', 'Greenock (Glasgow)', 'Inverness/Loch Ness',
              'Lerwick/Shetland', 'Liverpool',
              'Waterford (Dunmore E.)']

    for element in baltic:
        for p in ports:
            if p in element or element in p:
                return ['Baltic', 'E']
            elif ports[0] in element or element in ports[0]:
                return ['Baltic', 'E']

    for element in eastern_med:
        for p in ports:
            if p in element or element in p:
                return ['Eastern Med', 'E']

    for element in west_med:
        for p in ports:
            if p in element or element in p:
                return ['Western Med', 'E']
    for element in europe:
        for p in ports:
            if p in element or element in p:
                return ['Europe', 'E']

    return ['Carib', 'C']


def parse(rslt):
    sail_date = convert_date(rslt['date']['day'], rslt['date']['month'], rslt['date']['year'])
    number_of_nights = rslt['nights']
    return_date = calculate_days(sail_date, number_of_nights)
    brochure_name = rslt['title']
    if ' POST' in brochure_name or ' PRE ' in brochure_name or 'PACKAGE' in brochure_name:

        return
    else:
        pass
    cruise_id = '99'
    vessel_name = rslt['ship']['name']
    destination = get_destination(rslt['destination']['name'])
    if destination is None:
        dest_code = 'ANY'
        dest_name = 'DESTINATION'
    else:
        dest_code = destination[1]
        dest_name = destination[0]
    cruise_line_name = 'Azamra Club Cruises'
    package_id = ''
    if vessel_name == 'Azamara Journey':
        vessel_id = '408'
    else:
        vessel_id = '437'
    price_url = "https://www.azamaraclubcruises.com" + rslt['voyage_link']
    price_source = session.get(price_url).text
    price_page = BeautifulSoup(price_source, 'lxml')
    prices = price_page.find('div', {'class': 'voyage-pricing-section'})
    children = set()
    try:
        children = prices.find_all("div", recursive=False)
        interior = ''
        oceanview = ''
        veranda = ''
        continent = ''
        spa = ''
        ocean = ''
        world = ''
        suite = ''
        for child in children:
            if 'voyage-pricing-terms' not in child['class']:
                try:
                    current_room = child.find('div', {'class': 'voyage-pricing-single__room-title'}).text
                    actual_price = child.find('span', {'class': 'pricing-large'}).text.replace('$', '').replace(' USD',
                                                                                                                '').replace(
                        ',', '').replace('*', '')
                    if "Interior" in current_room:
                        if "Sold Out" in actual_price:
                            interior = 'N/A'
                        else:
                            interior = actual_price
                    if "Oceanview" in current_room:
                        if "Sold Out" in actual_price:
                            oceanview = 'N/A'
                        else:
                            oceanview = actual_price
                    if "Veranda" in current_room:
                        if "Sold Out" in actual_price:
                            veranda = 'N/A'
                        else:
                            veranda = actual_price
                    if "Continent" in current_room:
                        if "Sold Out" in actual_price:
                            continent = 'N/A'
                        else:
                            continent = actual_price
                    if "Spa" in current_room:
                        if "Sold Out" in actual_price:
                            spa = 'N/A'
                        else:
                            spa = actual_price
                    if "Ocean Suite" in current_room:
                        if "Sold Out" in actual_price:
                            ocean = 'N/A'
                        else:
                            ocean = actual_price
                    if "World Owner Suite" in current_room:
                        if "Sold Out" in actual_price:
                            world = 'N/A'
                        else:
                            world = actual_price
                except AttributeError:
                    print(child)
            if continent == "N/A":
                if spa == "N/A":
                    if world == "N/A":
                        suite = "N/A"
                    else:
                        suite = world
                else:
                    suite = spa
            else:
                suite = continent
    except AttributeError:
        interior = 'N/A'
        oceanview = 'N/A'
        veranda = 'N/A'
        continent = 'N/A'
        spa = 'N/A'
        ocean = 'N/A'
        world = 'N/A'
        suite = 'N/A'

    ports = []
    for port in rslt['ports']:
        ports.append(port['title'].split(', ')[0].strip())
    for p in ports:
        if 'Panama Canal' in p:
            dest_name = 'Panama Canal'
            dest_code = 'T'
            break
    if dest_code == "C":
        dest = split_carib(ports, dest_name, dest_code)
        dest_code = dest[1]
        dest_name = dest[0]
        for p in ports:
            if 'Oranjestad' in p:
                dest_code = "C"
                dest_name = 'East Carib'
    if 'Mediterranean' in dest_name:
        dest = split_europe(ports, dest_name, dest_code)
        dest_code = dest[1]
        dest_name = dest[0]
    if 'Europe' in dest_name:
        dest = split_europe(ports, dest_name, dest_code)
        dest_code = dest[1]
        dest_name = dest[0]
    if dest_code == 'I':
        if "Japan" in brochure_name:
            dest_code = "O"
            dest_name = 'Exotics'
    if dest_name == 'Australia/New Zealand':
        dest = split_australia(ports)
        dest_code = dest[1]
        dest_name = dest[0]
    if dest_code == 'C' and dest_name == 'Carib':
        dest = check_carib(ports, dest_name, dest_code)
        dest_code = dest[1]
        dest_name = dest[0]
    final_ports = ''
    for p in ports:
        final_ports += str(p + ', ')
    final_ports = final_ports.strip()[:-1]
    temp = [dest_code, dest_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
            package_id,
            brochure_name, number_of_nights, sail_date, return_date,
            interior, oceanview, veranda, suite, final_ports]
    temp2 = [temp]
    print(temp)
    all_cruises.append(temp2)


pool.map(parse, results)
pool.close()
pool.join()
print(len(all_cruises))


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Azamara Club Cruises.xlsx'
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
    worksheet.set_column("P:P", 50)
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


write_file_to_excell(all_cruises)
input("Press any key to continue...")
