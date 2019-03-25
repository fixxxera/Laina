import datetime
import os
import math

import requests
import xlsxwriter
from bs4 import BeautifulSoup

session = requests.session()
# headers = {
#     "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
#     "Accept": "application/json, text/plain, */*",
#     "Accept-Language": "en-US,en;q=0.5",
#     "Connection": "keep-alive",
#     "Refer": "https://www.carnival.com/",
#     "Cookie": "s_vi=[CS]v1|2C92325D851D3507-6000191120002B97[CE]; utag_main=v_id:015c7502d7e4006202069aa345b800044001900900bd0$_sn:1$_ss:1$_pn:1%3Bexp-session$_st:1496613542883$ses_id:1496611739620%3Bexp-session$dc_visit:1$dc_event:1%3Bexp-session$dc_region:eu-central-1%3Bexp-session; optimizelyEndUserId=oeu1496611739753r0.8896192772234034; optimizelySegments=%7B%221887831728%22%3A%22false%22%2C%221888511398%22%3A%22direct%22%2C%221898132135%22%3A%22none%22%2C%221913541698%22%3A%22ff%22%7D; optimizelyBuckets=%7B%228410210627%22%3A%228408200311%22%2C%226939620018%22%3A%226934321168%22%2C%228407180704%22%3A%228404570565%22%2C%228405600653%22%3A%228397910589%22%2C%227827422680%22%3A%227843031052%22%2C%228413440188%22%3A%228413780031%22%7D; optimizelyPendingLogEvents=%5B%5D; sandy-session-id=ea2797589057e761; sandy-client-id=25dd69e589fe1cde; website#lang=en; ASP.NET_SessionId=k3od3o0mj01qlmkisk5bwfvr; _ga=GA1.2.2045567053.1496611740; _gid=GA1.2.2131515016.1496611740; _gat_mobifyTracker=1; BigIPServerCarnival=!ZD0ZbNpX3f0SqrhbtW16dxLpZShBZOqP1DfOx0Ky4CqxbGv5oRuxuho0uGA2kj2A3lBFTlnX4zeFSVo=; cclHeader=%257B%2522TravelAdvisory%2522%253A%2522TravelAdvisory1%253Dtrue%2522%252C%2522DomainNotification%2522%253A%2522false%2522%257D; _gat_tealium_0=1; s_pers=%20s_vnum%3D1498856400818%2526vn%253D1%7C1498856400818%3B%20s_fid%3D1E868BE0F9C7536E-14F6109221605E84%7C1654378151742%3B%20gpv_pn%3Dcarnival.com%253Ahome%7C1496613551743%3B%20s_invisit%3Dtrue%7C1496613551745%3B%20s_dslv%3D1496611751746%7C1591219751746%3B%20s_dslv_s%3DFirst%2520Visit%7C1496613551746%3B%20s_nvr%3D1496611751748-New%7C1499203751748%3B; s_sess=%20s_cc%3Dtrue%3B%20s_ppvl%3Dcarnival.com%25253Ahome%252C51%252C35%252C887%252C1920%252C244%252C1920%252C1080%252C1%252CP%3B%20s_sq%3Dcarnivalprodus%253D%252526pid%25253Dcarnival.com%2525253Ahome%252526pidt%25253D1%252526oid%25253Djavascript%2525253A%2525253B%252526ot%25253DA%3B%20s_ppv%3Dcarnival.com%25253Ahome%252C55%252C51%252C952%252C1920%252C423%252C1920%252C1080%252C1%252CP%3B; _gat_tealium_1=1; carnivaltracking=1; NSE_AB=control; CCL_ExistingSession=true; CCL_IsReturner=false; _ceg.s=or1l0g; _ceg.u=or1l0g; mobify.webpush.client_id=42143ef692a13b39; mobify.webpush.user_state=Supported; mobify.webpush.activeVisit=1; mobify.webpush.subscription_cache=eyJvcHRlZF9vdXQiOmZhbHNlLCJjbGllbnRfaWQiOiI0MjE0M2VmNjkyYTEzYjM5Iiwic3Vic2NyaXB0aW9uX3N0YXR1cyI6InVua25vd24iLCJhY3RpdmUiOmZhbHNlLCJmYWtlIjp0cnVlfQ=="
# }
codes = ['']


def get_proxy():
    print("Looking for a working proxy server")
    soup = BeautifulSoup(requests.get('https://www.us-proxy.org').text, 'lxml')
    table = soup.find('table', {'id': 'proxylisttable'})
    tbody = table.find('tbody')
    proxies = []
    for tr in tbody:
        columns = tr.find_all('td')
        if columns[2].text in 'US' and columns[4].text in 'anonymous' and columns[6].text in 'yes':
            proxies.append("https://" + columns[0].text + ":" + columns[1].text)
    for p in proxies:
        try:
            proxy_line = {'https': p}
            resp = requests.get('https://www.carnival.com/cruisesearch/api/search?dest=a,bh,bm,c,e,et,h,m,q,t&numAdults=2&pageNumber=1&pageSize=8&showBest=true&sort=FromPrice&useSuggestions=true', proxies=proxy_line, timeout=10)
            if resp.ok:
                print("Found one!")
                return proxy_line
            else:
                print(p, "Not working")
        except requests.exceptions.ProxyError:
            print(p, "Not working")
        except requests.exceptions.ConnectTimeout:
            print(p, "Not working")
        except requests.exceptions.ReadTimeout:
            print(p, "Not working")


proxy = get_proxy()


def get_total_pages():
    count_url = "https://www.carnival.com/cruisesearch/api/search?dest=a,bh,bm,c,e,et,h,m,q,t&numAdults=2&pageNumber=1&pageSize=8&showBest=true&sort=FromPrice&useSuggestions=true"
    count_page = session.get(count_url, proxies=proxy)
    root = count_page.json()
    results = root['results']
    return results['totalResults']


def preformated(unformated):
    splitter = unformated.split('-')
    day = splitter[2]
    month = splitter[1]
    year = splitter[0]
    final_date = '%s/%s/%s' % (month, day, year)
    return final_date


itineraries = []
all_sailings = []
all_ports = set()
sailing_codes = []


def get_destination(param):
    if param == 'A':
        return ['Alaska', 'A']
    elif param == 'BH':
        return ['Bahamas', 'BH']
    elif param == 'BM':
        return ['Bermuda', 'BM']
    elif param == 'NN':
        return ['Canada & New England', 'NN']
    elif param == 'C':
        return ['Caribbean', 'C']
    elif param == 'E':
        return ['Europe', 'E']
    elif param == 'ET':
        return ['Transatlantic', 'X']
    elif param == 'H':
        return ['Hawaii', 'H']
    elif param == 'M':
        return ['Mexico', 'M']
    elif param == 'Q':
        return ['Cuba', 'C']
    elif param == 'T':
        return ['Panama Canal', 'T']
    elif param == "MB":
        return ['Baja Mexico', 'M']
    elif param == "CW":
        return ["West Carib", "C"]
    elif param == "CS":
        return ["Carib", "C"]
    elif param == "CE":
        return ["East Carib", "C"]
    elif param == "MR":
        return ["Mexico", "M"]


def get_vessel_id(vessel_na):
    if vessel_na == "Carnival Conquest":
        return "405"
    if vessel_na == "Carnival Sunshine":
        return "808"
    if vessel_na == "Carnival Glory":
        return "416"
    if vessel_na == "Carnival Legend":
        return "406"
    if vessel_na == "Carnival Miracle":
        return "426"
    if vessel_na == "Carnival Pride":
        return "398"
    if vessel_na == "Carnival Spirit":
        return "29"
    if vessel_na == "Carnival Triumph":
        return "30"
    if vessel_na == "Carnival Valor":
        return "436"
    if vessel_na == "Carnival Victory":
        return "31"
    if vessel_na == "Carnival Celebration":
        return "33"
    if vessel_na == "Carnival Ecstasy":
        return "34"
    if vessel_na == "Carnival Elation":
        return "35"
    if vessel_na == "Carnival Fantasy":
        return "37"
    if vessel_na == "Carnival Fascination":
        return "37"
    if vessel_na == "Carnival Holiday":
        return "39"
    if vessel_na == "Carnival Imagination":
        return "40"
    if vessel_na == "Carnival Inspiration":
        return "41"
    if vessel_na == "Carnival Jubilee":
        return "683"
    if vessel_na == "Carnival Paradise":
        return "45"
    if vessel_na == "Carnival Sensation":
        return "46"
    if vessel_na == "Carnival Liberty":
        return "441"
    if vessel_na == "Carnival Freedom":
        return "555"
    if vessel_na == "Carnival Splendor":
        return "662"
    if vessel_na == "Carnival Dream":
        return "694"
    if vessel_na == "Carnival Magic":
        return "724"
    if vessel_na == "Carnival Breeze":
        return "739"
    if vessel_na == "Carnival Sunshine":
        return "808"
    if vessel_na == "Carnival Vista":
        return "930"
    if vessel_na == "Carnival Horizon":
        return "28"


ids = set()


def split_carib(ports):
    bm = ['Kings Wharf, Bermuda']
    cu = ['Santiago de Cuba', 'Cienfuegos', 'Havana']
    wc = ['Costa Maya', 'Cozumel', 'Falmouth, Jamaica', 'George Town, Grand Cayman',
          'Ocho Rios']

    ec = ['Basseterre, St. Kitts', 'Bridgetown', 'Castries', 'Charlotte Amalie, St. Thomas',
          'Fort De France', 'Kingstown, St. Vincent', 'Philipsburg', 'Ponce, Puerto Rico',
          'Punta Cana, Dominican Rep', 'Roseau', 'San Juan', 'St. Croix, U.S.V.I.',
          "St. George's", "St. John's", 'Tortola, B.V.I']

    result = []
    iscu = False
    isec = False
    iswc = False
    isbm = False
    ports_list = []
    for i in range(len(ports)):
        if i == 0:
            ports_list.append(ports[i])
        else:
            ports_list.append(ports[i])
    for element in bm:
        for p in ports_list:
            if p in element:
                isbm = True
    if not isbm:
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

    if isbm:
        result.append("Bermuda")
        result.append("BM")
        # result.append("CU")
        return result
    elif iscu:
        result.append("Cuba")
        result.append("C")
        # result.append("CU")
        return result
    elif iswc:
        result.append("West Carib")
        result.append("C")
        # result.append("WC")
        return result
    elif isec:
        result.append("East Carib")
        result.append("C")
        # result.append("EC")
        return result
    else:
        result.append("Carib")
        result.append("C")
        # result.append("")
        return result


for code in codes:
    limit = get_total_pages()
    limit = int(math.ceil(limit/8))
    start_page = 1
    current_page = 1
    while current_page <= limit:
        print("Page", current_page, "of", limit)
        url = "https://www.carnival.com/cruisesearch/api/search?dest=a,bh,bm,c,e,et,h,m,q,t&numAdults=2&pageNumber=" + str(
            current_page) + "&pageSize=8&showBest=false&sort=FromPrice&useSuggestions=true"
        page = session.get(url, proxies=proxy)
        cruise_results = page.json()['results']
        for line in cruise_results['itineraries']:
            itineraries.append(line)
        current_page += 1
    print(len(itineraries))
    for line in itineraries:
        if line['id'] == 'GLB_SEA_LE_8_Mon':
            print('found')
        ports = []
        for p in line['ports']:
            ports.append(p['label'])
        port_string = ''
        for port in ports:
            port_string += ", " + port
        brochure_name = line['regionName'] + " from " + line['departurePortName']
        cruise_line_name = "Carnival US"
        vessel_name = line['shipName']
        number_of_nights = line['dur']
        destination_src = line['regionCode']
        destination = get_destination(destination_src)
        if 'AG' in destination_src:
            destination_code = 'ANY'
            destination_name = 'DESTINATION'
        else:
            destination_name = destination[0]
            destination_code = destination[1]
        if destination_code == 'C':
            destination = split_carib(ports)
            destination_name = destination[0]
        if destination_name == "Carib":
            if "Western" in brochure_name:
                destination_name = "West Carib"
            elif "Eastern" in brochure_name:
                destination_name = "East Carib"
        ids.add(vessel_name)
        vessel_id = get_vessel_id(vessel_name)
        cruise_id = "2"
        itinerary_id = ""
        for sailing in line['sailings']:
            if sailing['sailingId'] in sailing_codes:
                print("match")
                continue
            else:
                sailing_codes.append(sailing['sailingId'])
            departure_split = sailing['departureDate'].split('T')[0]
            arrival_split = sailing['arrivalDate'].split('T')[0]
            sail_date = preformated(departure_split)
            return_date = preformated(arrival_split)
            rooms = sailing['rooms']
            interior_bucket_price = str(rooms['interior']['price']).split('.')[0]
            oceanview_bucket_price = str(rooms['oceanview']['price']).split('.')[0]
            balcony_bucket_price = str(rooms['balcony']['price']).split('.')[0]
            suite_bucket_price = str(rooms['suite']['price']).split('.')[0]
            if interior_bucket_price == "0":
                interior_bucket_price = 'N/A'
            if oceanview_bucket_price == "0":
                oceanview_bucket_price = 'N/A'
            if balcony_bucket_price == "0":
                balcony_bucket_price = 'N/A'
            if suite_bucket_price == "0":
                suite_bucket_price = 'N/A'
            temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
                    itinerary_id, brochure_name, number_of_nights, sail_date, return_date, interior_bucket_price,
                    oceanview_bucket_price, balcony_bucket_price, suite_bucket_price, port_string[1:]]
            all_sailings.append(temp)
            print(temp)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Carnival US.xlsx'
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
                try:
                    worksheet.write_string(row_count, column_count, en, centered)
                except TypeError:
                    print(ship_entry)
            if column_count == 2:
                try:
                    worksheet.write_string(row_count, column_count, en, centered)
                except TypeError:
                    print(ship_entry)
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

write_file_to_excell(all_sailings)
input("Press any key to continue...")
