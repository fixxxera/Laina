import datetime
import os

import requests
import xlsxwriter
from bs4 import BeautifulSoup

session = requests.session()
sources = []
all_cruises = []
codes = set()
headers = {
    'Host': 'www.pocruises.com',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:53.0) Gecko/20100101 Firefox/53.0',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate',
    'Referer': 'http://www.pocruises.com/',
    'X-NewRelic-ID': 'Ug8PU1NTGwAHVFJbBgY=',
    'X-Requested-With': 'XMLHttpRequest',
    'Cookie': 'ens_pocruisecookie=criteo; mt.v=2.1114950516.1488213291674; ASP.NET_SessionId=a4hbasot2o44cjbw0gl3cm00; userCountry=UK; ecos.dt=1488213298533; __adal_first_visit=1488213292201; __adal_session_start=1488213292201; __adal_last_visit=1488213292201; __utma=239534620.1102973662.1488213293.1488213293.1488213293.1; __utmb=239534620.3.9.1488213297160; __utmc=239534620; __utmz=239534620.1488213293.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); __utmt=1; __qca=P0-451846045-1488213293106; yieldify_stc=1; yieldify_st=1; yieldify_ujt=4; _sp_ses.85c5=*; _sp_id.85c5=4a8d1003-525c-480c-ae49-db9fd0cfa9ca.1488213294.1.1488213294.1488213294.2f73cbd3-a9d4-4aaf-a434-6b55c170409e; yieldify_sale_ts=1488213294198; yieldify_visit=1; yieldify_iv=1; yieldify_location=%257B%2522country%2522%253A%2522Bulgaria%2522%252C%2522region%2522%253A%2522Sofia-Capital%2522%252C%2522city%2522%253A%2522Sofia%2522%257D; cookiesDirective=1',
    'Connection': 'keep-alive',
    'Cache-Control': 'max-age=0',
    'Content-Type': 'application/json'
}
print("Retrieving destination codes")
page = session.get('http://www.pocruises.com/Templates/PandO/AJAX/AZ_CruiseFilterData.aspx', headers=headers)
page.encoding = 'utf-8-sig'
page = page.json()
regions = page['regionids']
print(regions)
for dest in regions:
    if dest == 'CA':
        continue
    print("Retrieving result page", 1, "of destination", dest)
    page = session.get(
        'http://www.pocruises.com/Templates/PandO/Ajax/AZ_AddCruiseResults.aspx?departuremonths=&durationlist=&regionids=' + dest + '&_=1492358487236').text
    soup = BeautifulSoup(page, 'lxml')
    sources.append([dest, soup])
    page_number = 2
    last_page = int(soup.find('div', {'class': 'resultsWrapper'})['data-pagecount'])
    while page_number <= last_page:
        print("Retrieving result page", page_number, "of destination", dest)
        page = session.get(
            'http://www.pocruises.com/Templates/PandO/Ajax/AZ_AddCruiseResults.aspx?departuremonths=&durationlist=&regionids=' + dest + '&pn=' + str(
                page_number) + '&_=1492358886327')
        soup = BeautifulSoup(page.text, 'lxml')
        sources.append([dest, soup])
        page_number += 1


def get_dates(raw):
    split = raw.split()
    # SAIL DATE
    if len(split) == 6:
        day = split[0]
        month = split[1]
        year = split[5]
    else:
        day = split[0]
        month = split[1]
        year = split[2]
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
    sail = '%s/%s/%s' % (month, day, year)
    # RETURN DATE
    if len(split) == 6:
        day = split[3]
        month = split[4]
        year = split[5]
    else:
        day = split[4]
        month = split[5]
        year = split[6]
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
    ret = '%s/%s/%s' % (month, day, year)
    return [sail, ret]


def get_destination_name(param):
    if param == '1':
        return 'Mediterranean'
    elif param == '6':
        return 'Iberia & Canary Islands'
    elif param == '28':
        return 'Western Europe'
    elif param == '21':
        return 'Scandinavia'
    elif param == '32':
        return 'British Isles'
    elif param == '4':
        return 'Baltic'
    elif param == '33':
        return 'Canada'
    elif param == '13':
        return 'USA'
    elif param == '3':
        return 'Caribbean'
    elif param == '14':
        return 'South America'
    elif param == '24':
        return 'Central America'
    elif param == '26':
        return 'South Pacific'
    elif param == '23':
        return 'Exotics'
    elif param == '25':
        return 'Middle East'
    elif param == '22':
        return 'Africa'
    elif param == '15':
        return 'Indian Ocean'
    pass


def get_if_from_name(vname):
    if vname == 'AURORA':
        return '196'
    elif vname == 'OCEANA':
        return '197'
    elif vname == 'ARCADIA':
        return '194'
    elif vname == 'VENTURA':
        return '457'
    elif vname == 'BRITANNIA':
        return '0'
    elif vname == 'AZURA':
        return '0'
    elif vname == 'ADONIA':
        return '710'
    elif vname == 'ORIANA':
        return '198'


for source in sources:
    sections = source[1].find_all('section')
    for section in sections:
        dest_name = get_destination_name(source[0])
        cruise_id = '24'
        package_id = ''
        title = section.find('h2').text
        htmlspaced = title.replace('\r\n', ' ').strip()
        brochure_name = htmlspaced.split("   ")[0].strip()
        cruise_line_name = 'P&O Cruises UK'
        divs = section.find_all('div', recursive=False)
        infos_block = divs[2]
        prices_block = divs[1]
        infos = infos_block.find_all('div', recursive=False)[1].find_all('div', recursive=False)[0]
        raw_dates = infos.find('h3').text
        dates = get_dates(raw_dates)
        sail_date = dates[0]
        return_date = dates[1]
        number_of_nights = section.find('span').text.split(" NIGHTS")[0].strip()
        meta = infos.find('h5').text.split()
        vessel_name = meta[0]
        vessel_id = get_if_from_name(vessel_name)
        sailing_code = meta[1].replace('(', '').replace(')', '')
        price_rows = prices_block.find('ul').find_all('li')
        interior_bucket_price = ''
        oceanview_bucket_price = ''
        balcony_bucket_price = ''
        suite_bucket_price = ''
        if sailing_code in codes:
            continue
        else:
            codes.add(sailing_code)
        for p in price_rows:
            siblings = p.find('a', recursive=False)
            if len(siblings.findChildren()) == 4:
                htmlspaced = siblings.find('h3').text.replace('\r\n', ' ').strip()
                fixed = ([htmlspaced.split("  ")[0], htmlspaced.split("  ")[len(htmlspaced.split("  ")) - 1]])
                if fixed[0] == 'INSIDE CABINS':
                    interior_bucket_price = 'N/A'
                elif fixed[0] == 'OUTSIDE CABINS':
                    oceanview_bucket_price = "N/A"
                elif fixed[0] == 'BALCONY CABINS':
                    balcony_bucket_price = "N/A"
                elif fixed[0] == 'MINI SUITES':
                    suite_bucket_price = "N/A"
            elif len(siblings.findChildren()) == 7:
                room_type = siblings.find('h3').text.strip()
                price_element = siblings.findChildren()[3].text.split()[1].replace('£', '').replace('pp', '').replace(
                    ',', '')
                if price_element == '':
                    price_element = "N/A"
                if room_type == 'INSIDE CABINS':
                    interior_bucket_price = price_element
                elif room_type == 'OUTSIDE CABINS':
                    oceanview_bucket_price = price_element
                elif room_type == 'BALCONY CABINS':
                    balcony_bucket_price = price_element
                elif room_type == 'MINI SUITES':
                    suite_bucket_price = price_element
        if interior_bucket_price == '':
            interior_bucket_price = "N/A"
        if oceanview_bucket_price == '':
            oceanview_bucket_price = 'N/A'
        if balcony_bucket_price == '':
            balcony_bucket_price = 'N/A'
        if suite_bucket_price == '':
            suite_bucket_price = 'N/A'
        temp = [dest_name, dest_name, vessel_id, vessel_name, cruise_id, cruise_line_name, package_id, brochure_name,
                number_of_nights, sail_date, return_date, interior_bucket_price, oceanview_bucket_price,
                balcony_bucket_price, suite_bucket_price]
        print(temp)
        temp2 = [temp]
        all_cruises.append(temp2)


def write_file_to_excell(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- P&O UK.xlsx'
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


write_file_to_excell(all_cruises)
input("Press any key to continue...")
