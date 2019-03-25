import datetime
import os
from multiprocessing.dummy import Pool as ThreadPool

import requests
import xlsxwriter

session = requests.session()
pool = ThreadPool(10)
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:64.0) Gecko/20100101 Firefox/64.0",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Accept-Encoding": "gzip, deflate, br",
    "Host": "www.costacruises.com",
    "Refer": "https://www.costacruises.com/cruises.html",
    "Origin": "https://www.costacruises.com",
    "Cookie": "countryCode=US; continentCode=NA; CostaPersistent=New; ak_bmsc=2BB354C69D92C971C748437877BB90184D5C892FD06A0000FC5FCF5B764F2D43~pldZXWuaM7qmq0Mg+wbGkxgiy+Tl2IQvIzM4j+7w/DgHuaf3gCxvRuI9e+6tqEVQbpUxhAofK2BIpMOfKzldSNhErwL/1UntY2rvrQI5t8wc3wFlPqZvYGaaL2NYisrjKbczmFlFL3O4t6xyhu40C0uIxrlmvZoQxwYzW4ab7WP7v81OqziRt0Zgb92x4Zn4+QpsrDJgqLH2aDEdCUDZha6tqH8u9Q/PjXmHz0SCb41+bRyhZNzdOQGCU9yQVOA6DR9/Y39b8iuc/cDGQ6iZQ/dpcXf4wXzErWI+qp2Ep0tor4bQ3Xy0pkNzrjABuILYNJ; _abck=0C45917816B45B87B515071D503CDC284D5C892FD06A0000FC5FCF5B71FA5C52~-1~RRTeNxMbZ9QFkCQ7q5rLqB9mAGNTFPdSSKlpcPUBmy4=~-1~-1; bm_sz=8522CB7E9E16640C42C21E5FFDA93AFF~QAAQL4lcTZLkNThmAQAA9/IOolvdeX/nkdwvJGKWnMOC9ZSp7NltvfitAjSgiyWmk2FdfaK9X6HN7oQBhgwCYp9Q4q1+tPaFqxQIvNJpy+orNfAPtDtzw0JTJRqXIXIdp2TTSD176/HR1r8eRdnqmNFXWVn43dVIRY+5kihF5oDAKl2HZ9XL88IoX03HXcE4hUv/I5Y=; bm_mi=DEA8105810EC3D83C8263D0C44EB78E5~myj7MpFwV6j1PVzrh80VOruGbFNhOXP4/LzotUy8S3nWom1x4zNooWw9zV1LjlNS75eyZHXsZxiQ9wkMBV4tHShBpe6vWvn0zM1vX9XE6IwIBfwJaMf6koe8JwKCBV9ZJS3F8asSp3ndetz6kva+8gHo2SBA0gptb/nJ+2Yk4+l1ekSSG5ulU0aeFuC1NFLVNGs0+l994fL5Fw9YhPtj/1vPmDNDVHnRObMjdRKTt8atq/eAWdprnbB5DOIk3IzXWNwpl3qQ+brgUbPckuFo+A==; bm_sv=6DDEA657C01FF2162FA5D66A6739702D~mKziNvtqJeDZ1JojfBq7QyRCo0Gf7a5FVAEkS3YdlSfPUnvYUyM/XhVBIcaRKurcqYnmdh5euWRmqXbKlfr5O+G0HUDTH7tm3NEXZYhUthQzb+c975AX/j/8qmElB5ujke3VZUVRbsEfK36BESavcdjMtYc/PcmeW/L2v/bDLl0=; ref_c=http://www.costacruise.com/; languageCode=en_US",
    "loyaltynumber": "",
    "TE": "Trailers"
}
price_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:64.0) Gecko/20100101 Firefox/64.0",
    "Accept": "application/json",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate, br",
    "Refer": "https://www.costacruises.com/cruises.html",
    "Origin": "https://www.costacruises.com",
    "country": "US",
    "Cookie": "countryCode=US; continentCode=NA; CostaPersistent=New; ak_bmsc=2BB354C69D92C971C748437877BB90184D5C892FD06A0000FC5FCF5B764F2D43~pldZXWuaM7qmq0Mg+wbGkxgiy+Tl2IQvIzM4j+7w/DgHuaf3gCxvRuI9e+6tqEVQbpUxhAofK2BIpMOfKzldSNhErwL/1UntY2rvrQI5t8wc3wFlPqZvYGaaL2NYisrjKbczmFlFL3O4t6xyhu40C0uIxrlmvZoQxwYzW4ab7WP7v81OqziRt0Zgb92x4Zn4+QpsrDJgqLH2aDEdCUDZha6tqH8u9Q/PjXmHz0SCb41+bRyhZNzdOQGCU9yQVOA6DR9/Y39b8iuc/cDGQ6iZQ/dpcXf4wXzErWI+qp2Ep0tor4bQ3Xy0pkNzrjABuILYNJ; _abck=0C45917816B45B87B515071D503CDC284D5C892FD06A0000FC5FCF5B71FA5C52~-1~RRTeNxMbZ9QFkCQ7q5rLqB9mAGNTFPdSSKlpcPUBmy4=~-1~-1; bm_sz=8522CB7E9E16640C42C21E5FFDA93AFF~QAAQL4lcTZLkNThmAQAA9/IOolvdeX/nkdwvJGKWnMOC9ZSp7NltvfitAjSgiyWmk2FdfaK9X6HN7oQBhgwCYp9Q4q1+tPaFqxQIvNJpy+orNfAPtDtzw0JTJRqXIXIdp2TTSD176/HR1r8eRdnqmNFXWVn43dVIRY+5kihF5oDAKl2HZ9XL88IoX03HXcE4hUv/I5Y=; bm_mi=DEA8105810EC3D83C8263D0C44EB78E5~myj7MpFwV6j1PVzrh80VOruGbFNhOXP4/LzotUy8S3nWom1x4zNooWw9zV1LjlNS75eyZHXsZxiQ9wkMBV4tHShBpe6vWvn0zM1vX9XE6IwIBfwJaMf6koe8JwKCBV9ZJS3F8asSp3ndetz6kva+8gHo2SBA0gptb/nJ+2Yk4+l1ekSSG5ulU0aeFuC1NFLVNGs0+l994fL5Fw9YhPtj/1vPmDNDVHnRObMjdRKTt8atq/eAWdprnbB5DOIk3IzXWNwpl3qQ+brgUbPckuFo+A==; bm_sv=6DDEA657C01FF2162FA5D66A6739702D~mKziNvtqJeDZ1JojfBq7QyRCo0Gf7a5FVAEkS3YdlSfPUnvYUyM/XhVBIcaRKurcqYnmdh5euWRmqXbKlfr5O+G0HUDTH7tm3NEXZYhUthQzb+c975AX/j/8qmElB5ujke3VZUVRbsEfK36BESavcdjMtYc/PcmeW/L2v/bDLl0=; ref_c=http://www.costacruise.com/; languageCode=en_US; cruiseIds=MG14181213_FDF10029; AWSELB=3B2117451A1D761F72F9EA261034552E59B1034A7208D1089F1F64D37F544F9D0AE5D9E7E207C196AF44A30A9C33A5FF59D68847AD6A6871886F7CA055A1605BD1F8B9F8B8",
    "currencycode": "USD",
    "loyaltynumber": "",
    "brand": "costa",
    "locale": "en_US",
    "agency": "63238180"
}

urls = []
voyage_ids = []
voyages = []
to_write = []

# codes = ['CA', 'PG', 'PA', 'IO', 'ME', 'NO', 'SA', 'TR', 'RW']
codes = ['CA']

for code in codes:
    print("Downloading", "'" + code + "'", "destination")
    # not sold out
    page = requests.get(
        'https://www.costacruises.com/search/costa_en_US/cruisesearch?&start=0&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&fq={!tag=destinationTag}destinationIds:(\\' + code + ')&rows=5&fq=departDate:[NOW/DAY%2B2DAY%20TO%20*]',
        headers=headers)
    cruise_results = page.json()
    print(page.reason)
    total_results = int(cruise_results['results'])
    print("Found", total_results, "results")
    current_start = 0
    not_sold_out = 0
    sold_out = 0
    while current_start <= total_results:
        page = requests.get('https://www.costacruises.com/search/costa_en_US/cruisesearch?&start='+str(current_start)+'&fq=(soldOut:(false)%20AND%20price_USD_anonymous:[1%20TO%20*])&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&fq={!tag=destinationTag}destinationIds:(\\' + code + ')&rows=5&fq=departDate:[NOW/DAY%2B2DAY%20TO%20*]',
                            headers=headers)
        urls.append([code, page.json()['searchResults'], False])
        current_start += 5
    # Sold out
    page = requests.get('https://www.costacruises.com/search/costa_en_US/cruisesearch?&start='+str(current_start)+'&fq=(soldOut:(true)%20AND%20price_USD_anonymous:[1%20TO%20*])&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&fq={!tag=destinationTag}destinationIds:(\\' + code + ')&rows=5&fq=departDate:[NOW/DAY%2B2DAY%20TO%20*]',
                        headers=headers)
    cruise_results = page.json()
    print(page.reason)
    total_results = int(cruise_results['results'])
    current_start = 0
    not_sold_out = 0
    sold_out = 0
    while current_start <= total_results:
        page = requests.get('https://www.costacruises.com/search/costa_en_US/cruisesearch?&start='+str(current_start)+'&fq=(soldOut:(true)%20AND%20price_USD_anonymous:[1%20TO%20*])&sort=departDate%20asc,price_USD_anonymous%20asc&group.sort=departDate%20asc,price_USD_anonymous%20asc&fq={!tag=destinationTag}destinationIds:(\\' + code + ')&rows=5&fq=departDate:[NOW/DAY%2B2DAY%20TO%20*]',
                            headers=headers)
        urls.append([code, page.json()['searchResults'], True])
        current_start += 5


def convert_date(not_formatted):
    splitter = not_formatted.split("-")
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


def get_destination(param):
    if param == 'CA':
        return ['Caribbean', 'C']
    elif param == 'PG':
        return ['Dubai & UAE', 'SN']
    elif param == 'PA':
        return ['Far East', 'O']
    elif param == 'IO':
        return ['Indian Ocean', 'O']
    elif param == 'ME':
        return ['Mediterranean', 'MED']
    elif param == 'NO':
        return ['Europe', 'E']
    elif param == 'SA':
        return ['South America', 'S']
    elif param == 'TR':
        return ['Transoceanic', 'TO']
    elif param == 'RW':
        return ['World', 'W']
    else:
        return ["MISSING.............................................................", param]


def get_vessel_id(ves_name):
    if ves_name == "Insignia":
        return "429"
    if ves_name == "Marina":
        return "700"
    if ves_name == "Nautica":
        return "495"
    if ves_name == "Regatta":
        return "430"
    if ves_name == "Riviera":
        return "770"
    if ves_name == "Sirena":
        return "938"


def match_by_meta(ports_visited):
    bermuda = ['Hamilton', 'St. George']
    hawaii = ['Hilo', 'Honolulu', 'Kahului', 'Nawiliwili']
    panama_canal = ['Colon', 'Fuerte Amador']
    mexico = ['Acapulco', 'Huatulco', 'Cabo San Lucas']
    west_med = ['Ajaccio', 'Alicante', 'Almeria', 'Amalfi', 'Antibes', 'Arrecife', 'Bandol', 'Barcelona', 'Bastia',
                'Belfast', 'Biarritz', 'Bilbao', 'Bordeaux', 'Brest', 'Cagliari', 'Calvi', 'Cannes', 'Cartagena',
                'Casablanca', 'Catania', 'Cinque Terre', 'Cork', 'Corner Brook', 'Dublin', 'Florence', 'Funchal',
                'Gaeta', 'Gibraltar', 'Gijon', 'Huelva', 'Ibiza', 'La Coruna', 'La Rochelle',
                'Las Palmas de Gran Canaria', 'Lisbon', 'London', 'Lorient', 'Mahon', 'Malaga', 'Messina',
                'Monte Carlo', 'Montreal', 'Naples', 'Olbia', 'Oporto', 'Palamos', 'Palermo', 'Palma de Mallorca',
                'Paris', 'Porto Santo Stefano', 'Portofino', 'Port-Vendres', 'Provence', 'Ravenna', 'Rome', 'Roses',
                'Saint-Pierre', 'Saint-Tropez', 'Santa Cruz de La Palma', 'Santa Cruz de Tenerife',
                'Santiago de Compostela', 'Sete', 'Seville', 'Sorrento', 'St. Peter Port', 'Tangier', 'Taormina',
                'Toulon', 'Trois-Rivieres', 'Umbria', 'Valencia', 'Venice', 'Villefranche']
    east_med = ['Argostoli', 'Athens', 'Chania', 'Corfu', 'Dubrovnik', 'Gythion', 'Heraklion', 'Jerusalem', 'Katakolon',
                'Koper', 'Kotor', 'Limassol', 'Monemvasia', 'Mykonos', 'Patmos', 'Rhodes', 'Rijeka', 'Santorini',
                'Split', 'Thessaloniki', 'Tirana', 'Valletta', 'Volos', 'Zadar', 'Zakynthos']
    exotics = ['Aqaba', 'Dubai', 'Luxor', 'Muscat', 'Salalah', 'Sharm El Sheikh']
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in bermuda:
            return 'Bermuda'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in hawaii:
            return 'Hawaii'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in panama_canal:
            return "Panama Canal"
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in mexico:
            return 'Mexico'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in west_med:
            return 'West Mediterranean'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in east_med:
            return 'East Mediterranean'
    for i in range(1, len(ports_visited)):
        if ports_visited[i]['name'] in exotics:
            return 'Exotics'
    return 'Carib'


def split_carib(ports):
    cu = ['Santiago de Cuba', 'Cienfuegos', 'Havana']
    wc = ['Costa Maya', 'Cozumel', 'Falmouth, Jamaica', 'George Town, Grand Cayman',
          'Ocho Rios']

    ec = ['Basseterre, St. Kitts', 'Bridgetown', 'Castries', 'Charlotte Amalie, St. Thomas',
          'Fort De France', 'Kingstown, St. Vincent', 'Philipsburg', 'Ponce, Puerto Rico',
          'Punta Cana, Dominican Rep', 'Roseau', 'San Juan', 'St. Croix, U.S.V.I.',
          "St. George's", "St. John's", 'Tortola, B.V.I']

    bm = ['Kings Wharf, Bermuda']
    result = []
    iscu = False
    isec = False
    iswc = False
    ports_list = []
    for i in range(len(ports)):
        if i == 0:
            pass
        else:
            ports_list.append(ports[i]['name'])
    for element in cu:
        for p in ports_list:
            if p in element:
                iscu = True
    if not iscu:
        for element in wc:
            for p in ports_list:
                if p in element:
                    iswc = True
    if not iswc:
        for element in ec:
            for p in ports_list:
                if p in element:
                    isec = True
    if iscu:
        result.append("Cuba")
        result.append("C")
        result.append("CU")
        return result
    elif iswc:
        result.append("West Carib")
        result.append("C")
        result.append("WC")
        return result
    elif isec:
        result.append("East Carib")
        result.append("C")
        result.append("EC")
        return result
    else:
        result.append("Carib")
        result.append("C")
        result.append("")
        return result


counter = 0
codes = []


def parse(ur):
    for result in ur[1]:
        brochure_name = result['title']
        vessel_name = result['shipName']
        cruise_line_name = "Costa Cruises"
        number_of_nights = (result['duration'])
        itinerary = requests.get("https://www.costacruises.com/api/v2/price/itinerary/" + result['itineraryId'],
                                 headers=price_headers).json()
        print("https://www.costacruises.com/api/v2/price/itinerary/" + result['itineraryId'])
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
            sail_date = convert_date(sailing['arriveDate'])
            return_date = convert_date(sailing['departDate'])
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
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Costa Cruises.xlsx'
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
