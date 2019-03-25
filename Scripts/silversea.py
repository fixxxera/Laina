import datetime
import math
import os

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from multiprocessing.dummy import Pool as ThreadPool

pool = ThreadPool(5)


def preformat_date(unformated):
    splitter = unformated.split()
    day = splitter[1]
    if day[0] == '0':
        day = day[1]
    month = splitter[0]
    year = splitter[2]
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


def calculate_days(sail_date_param, number_of_nights_param):
    date = datetime.datetime.strptime(sail_date_param, "%m/%d/%Y")
    try:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param))
    except ValueError:
        calculated = date + datetime.timedelta(days=int(number_of_nights_param.split("-")[1]))
    calculated = calculated.strftime("%#m/%#d/%Y")
    return calculated


def get_destination(param):
    if param == 'Africa & Indian Ocean' or "Indian Ocean" in param or "Africa" in param:
        return ['Africa & Indian Ocean', 'AI']
    elif param == 'Alaska':
        return ['Alaska', 'A']
    elif param == 'Antarctica':
        return ['Antarctica', 'AN']
    elif param == 'Arctic And Greenland' or param == "Arctic & Greenland" or "Arctic" in param or "greenland" in param or "greeland" in param or 'Greeland' in param:
        return ['Arctic & Greenland', 'AG']
    elif param == 'Asia' or "ASIA" in param:
        return ['Asia & Pacific', 'O']
    elif param == 'Australia & New Zealand' or "Australia" in param or 'NEW ZEALAND' in param or 'New Zealand' in param:
        return ['Australia/New Zealand & S.Pacific', 'P']
    elif param == 'Canada & New England':
        return ['Canada/New England', 'N']
    elif param == 'Caribbean & Central America' or "Caribbean" in param or "Central America" in param:
        return ['Caribbean', 'C']
    elif param == 'Galápagos Islands' or "G A L Á P A G O S " in param or "GALÁPAGOS ISLANDS" in param:
        return ['Glapagos', 'GP']
    elif param == 'Mediterranean' or "M E D I T E R R A N E A N" in param or "MEDITERRANEAN" in param:
        return ['Med', 'E']
    elif param == 'Northern Europe & British Isles' or "Northern Europe" in param or "British Isles" in param:
        return ['Europe', 'E']
    elif param == 'Russian Far East':
        return ['Russian Far East', 'R']
    elif param == 'South America' or 'AMERICAN WEST COAST' in param or 'American West Coast' in param:
        return ['South America & Antarctica', 'S']
    elif param == 'South Pacific Islands':
        return ['South America & Antarctica', 'T']
    elif param == 'Transoceanic':
        return ['Transoceanic', 'DK']
    elif param == 'World Cruises':
        return ['World', 'W']
    elif param == 'Grand Voyages':
        return ['GrandVoyages', 'G']
    else:
        print(param)


init_response = requests.get('https://www.silversea.com/find-a-cruise/cruise-results.html')
init_soup = BeautifulSoup(init_response.text, 'lxml')
total_voyages = init_soup.find('div', {'id': 'v2-matching-value'}).text
pages = int(math.ceil(int(total_voyages) / 8))
urls = []
all_cruises = []
for page in range(1, pages):
    urls.append(
        'https://www.silversea.com/find-a-cruise/cruise-results/_jcr_content/par/findyourcruise.resultsv2.destination_all.date_all.duration_all.ship_all.cruisetype_all.port_all.psize_9.features_all.page_' + str(
            page) + '.html')


def parse(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    result_pane = soup.find('div', {'class': 'clearfix c-fyc-v2__results__wrapper'})
    results = result_pane.findAll('div', {'class': 'c-fyc-v2__result'})
    for result in results:
        ship_name = result.find('li', {'class': 'cruise-ship_2 hidden-xs'}).text.strip()
        title = result.find('li', {'class': 'cruise-destination'}).text.replace('\n', ' ').strip().title()
        sail_date = preformat_date(
            result.find('dd', {'class': 'c-fyc-v2__result__content__summary__depdate'}).text.strip())
        duration = result.find('li', {
            'class': 'c-fyc-v2__result__content__summary__item c-fyc-v2__result__content__summary__item__duration'}).find(
            'dd').text.split()[0].strip()

        ports_wrapper = result.find('li', {'class': 'destination-ports'})
        ports_tag = ports_wrapper.findAll('span')[1:]
        ports = []
        for tag in ports_tag:
            port_name = tag.text.strip().title() + ', '
            if port_name == "Day At Sea, ":
                pass
            else:
                ports.append(port_name)
        if len(ports) == 0:
            ports.append("No ports published in site  ")
        if title == 'To':
            title = "No title published in site"
        return_date = calculate_days(sail_date, duration)
        itin_url = 'https://www.silversea.com' + result.find('a', {'class': 'c-fyc-v2__result__link'})['href'].strip()
        details_response = requests.get(itin_url)
        soup = BeautifulSoup(details_response.text, 'lxml')
        try:
            destination_raw = soup.find('h2', {'class': 'cruise_header__subtitle'}).text.strip().split()
        except AttributeError:
            destination_raw = soup.find('div', {'class': 'cruise-2018-overview-description-cruise'}).text.strip().split()
        destination = []
        for word in destination_raw:
            if "cruise" in word.strip() or "Cruise" in word.strip() or "expedition" in word.strip() or "Expedition" in word.strip():
                break
            else:
                destination.append(word)
                destination.append(" ")
        destination = ''.join(destination)
        destination_set = get_destination(destination.title().strip())
        dest_code = destination_set[1]
        dest_name = destination_set[0]
        price_grid = soup.find('div', {'class': 'row c-suitelist__row-card'})
        try:
            rooms = price_grid.findAll('div', {'data-target': '.bs-modal-lg'}, recursive=False)
        except AttributeError:
            price_grid = soup.find('div', {'class': 'cruise-2018-suites-fares-main row'})
            rooms = price_grid.findAll('div', {'class': 'cruise-2018-suites-fares-container'})
        suite_price = 'N/A'
        balcony = 'N/A'
        oceanview = 'N/A'
        if ship_name in "Silver Cloud Expedition":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                    # tag = room.find('span', {
                    #     'class': 'cruise-2018-suites-fares-description-title'}).text
                    # if tag == "Vista Suite":
                    #     prices.append(
                    #         int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                    #                 1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Veranda Suite" or tag == "Veranda Deluxe":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Medallion Suite" or tag == "Silver Suite" or tag == "Royal Suite" or tag == "Grand Suite" or tag == "Owner Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Discoverer":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite" or tag == "View Suite" or tag == "Explorer Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Veranda Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Medallion Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Explorer":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite" or tag == "View Suite" or tag == "Explorer Suite" or tag == "Adventurer Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Veranda Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Silver Suite" or tag == "Owners Suite" or tag == "Medallion Suite" or tag == "Grand Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Galapagos":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Explorer Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Veranda Suite" or tag == "Deluxe Veranda Suite" or tag == "Terrace Veranda":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Silver Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Muse":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite" or tag == "Panorama Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Superior Veranda Suite" or tag == "Deluxe Veranda Suite" or tag == "Classic Veranda Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Silver Suite" or tag == "Royal Suite" or tag == "Grand Suite" or tag == "Owner Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Shadow":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Veranda Suite" or tag == "Terrace Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Silver Suite" or tag == "Royal Suite" or tag == "Grand Suite" or tag == "Owner Suite" or tag == "Medallion Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Spirit":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite" or tag == "Panorama":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Superior Veranda Suite" or tag == "Deluxe Veranda Suite" or tag == "Classic Veranda Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Silver Suite" or tag == "Royal Suite" or tag == "Grand Suite" or tag == "Owner Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Whisper":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Veranda Suite" or tag == "Terrace Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Silver Suite" or tag == "Royal Suite" or tag == "Grand Suite" or tag == "Owner Suite" or tag == "Medallion Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        elif ship_name in "Silver Wind":
            # OCEANVIEW
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Vista Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                oceanview = min(prices)
            except ValueError:
                pass
            # BALCONY
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Veranda Suite" or tag == "Midship Veranda Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                balcony = min(prices)
            except ValueError:
                pass
            # SUITE
            prices = []
            for room in rooms:
                try:
                    tag = room.find('span', {
                        'class': 'cruise-2018-suites-fares-description-title'}).text
                    if tag == "Silver Suite" or tag == "Royal Suite" or tag == "Grand Suite" or tag == "Owner Suite" or tag == "Medallion Suite":
                        prices.append(
                            int(room.find('span', {'class': 'cruise-2018-suites-fares-price-text'}).text.split()[
                                    1].replace(',', '')))
                except AttributeError:
                    continue
            try:
                suite_price = min(prices)
            except ValueError:
                pass
        temp = [dest_code, dest_name, '', ship_name, '', "SilverSea", '',
                title, duration, sail_date, return_date,
                'N/A', oceanview, balcony, suite_price, (''.join(ports))[:-2]]
        print(temp)
        temp2 = [temp]
        all_cruises.append(temp2)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- SilverSea.xlsx'
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
    red = workbook.add_format({'bold': True})
    red.set_font_color('#FF0000')
    red.set_bold(True)
    red.set_bg_color('#FFFF00')
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
                        if 'No ports published in site' in str(r):
                            worksheet.write_string(row_count, column_count, str(r), red)
                        else:
                            worksheet.write_string(row_count, column_count, str(r), centered)
                        column_count += 1
                except ValueError:
                    worksheet.write_string(row_count, column_count, r, centered)
                    column_count += 1

            row_count += 1
    workbook.close()
    pass


pool.map(parse, urls)
pool.close()
pool.join()
write_file_to_excell(all_cruises)
print(len(all_cruises))
