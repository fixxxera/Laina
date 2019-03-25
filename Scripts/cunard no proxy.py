import datetime
import os
from json import JSONDecodeError

import math
import requests
from multiprocessing.dummy import Pool as ThreadPool

import xlsxwriter
from bs4 import BeautifulSoup

session = requests.session()
a = requests.adapters.HTTPAdapter(max_retries=30)
session.mount('http://', a)
pool = ThreadPool(1)
pool2 = ThreadPool(1)
pool3 = ThreadPool(1)
to_write = []
headers = {
    'Host': 'www.cunard.com',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:54.0) Gecko/20100101 Firefox/54.0',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive'
}


# def get_proxy():
#     print("Looking for a working proxy server")
#     soup = BeautifulSoup(requests.get('https://www.us-proxy.org').text, 'lxml')
#     table = soup.find('table', {'id': 'proxylisttable'})
#     tbody = table.find('tbody')
#     proxies = []
#     for tr in tbody:
#         columns = tr.find_all('td')
#         if columns[2].text in 'US' and columns[4].text in 'anonymous' and columns[6].text in 'no':
#             proxies.append("http://" + columns[0].text + ":" + columns[1].text)
#     for p in proxies:
#         try:
#             proxy_line = {'http': p}
#             resp = requests.get('http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=AFRI&pg=1&rt=&sort=', proxies=proxy_line, timeout=10)
#             if resp.ok:
#                 print("Found one!")
#                 return proxy_line
#             else:
#                 print(p, "Not working")
#         except requests.exceptions.ProxyError:
#             print(p, "Not working")
#         except requests.exceptions.ConnectTimeout:
#             print(p, "Not working")
#         except requests.exceptions.ReadTimeout:
#             print(p, "Not working")
#         # ahhaha


# proxies = get_proxy()
destinations = ['AFRI', 'ATIS', 'AUST', 'BRIT', 'CA', 'CARI', 'CEAM', 'FARE', 'HAWA', 'INOC', 'MEDI', 'MIDE', 'NYSO',
                'NOEU', 'SCAN', 'SOAM', 'SONY', 'BALT', 'TRAN', 'USA', 'NOAM', 'WORL']

names = []
major_urls = []
failed_urls = []


def convert_date(unformatted, ye):
    splitter = unformatted.split()
    day = splitter[0]
    month = splitter[1]
    year = ye
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


def convert_return(unformatted):
    splitter = unformatted.split()
    day = splitter[0]
    month = splitter[1]
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
    pass


def parse(de):
    try:
        url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=1&rt=&sort='
        count = int(session.get(url=url, headers=headers).json()['Count'])
        pages = math.ceil(count / 5)

    except JSONDecodeError:
        try:
            url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=1&rt=&sort=&ajax=1'
            count = int(session.get(url=url).json()['Count'])
            pages = math.ceil(count / 5)
        except JSONDecodeError:
            print("skipped destination url:", de)
            major_urls.append(de)
            return
    index = 1
    while index <= int(pages):
        try:
            url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=' + str(
                index) + '&rt=&sort='
            page = session.get(url=url).json()['Html']
        except JSONDecodeError:
            try:
                url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=' + str(
                    index) + '&rt=&sort=&ajax=1'
                page = session.get(url=url).json()['Html']
            except JSONDecodeError:
                print("skipped result page:", url)
                return
        soup = BeautifulSoup(page, 'lxml')
        results_on_page = soup.find_all('li', {'class': 'resultRow'})
        for result in results_on_page:
            brochure_name = result.find('h2').text
            vessel_name = result.find('ul', {'class': 'resultOverview'}).find('img')['alt']
            result_overview = result.find('ul', {'class': 'resultOverview'})
            duration = result_overview.contents[1].text.split()[0]
            dates = result_overview.contents[2].text.split(" - ")
            yr = dates[1].split()[2]
            sail_date = convert_date(dates[0], yr)
            return_date = convert_return(dates[1])
            button_url = ''
            itinerary_id = ''
            cruise_id = '6'
            if vessel_name == "Queen Elizabeth":
                vessel_id = '512'
            elif vessel_name == "Queen Mary 2":
                vessel_id = '53'
            elif vessel_name == "Queen Victoria":
                vessel_id = '188'
            cruise_line_name = 'Cunard Cruises'
            try:
                buttons = result.find('div', {'class': 'pricing'}).find_all('a')
                if len(buttons) > 1:
                    button_url = buttons[1]['href']
                else:
                    button_url = buttons[0]['href']
            except KeyError:
                print('Failed:', vessel_name, dates)
            if 'http' in button_url:
                inside = "N/A"
                oceanview = "N/A"
                balcony = "N/A"
                suite = "N/A"
                temp = [de, de, vessel_id, vessel_name, cruise_id, cruise_line_name,
                        itinerary_id, brochure_name, duration, sail_date, return_date, inside,
                        oceanview, balcony, suite]
                if [vessel_name, sail_date, return_date] in names:
                    continue
                else:
                    names.append([vessel_name, sail_date, return_date])
                    to_write.append(temp)
                    print(temp)
            else:
                button_url = 'http://www.cunard.com' + button_url
                prices_page = session.get(url=button_url).text
                soup = BeautifulSoup(prices_page, 'lxml')
                gradelist = soup.find('ul', {'class': 'gradeList'})
                try:
                    lis = gradelist.find_all('li')
                    inside = "N/A"
                    oceanview = "N/A"
                    balcony = "N/A"
                    suite = "N/A"
                    for l in lis:
                        navlink = l.find('span')
                        room = navlink.find('strong').text
                        price = navlink.find('em').text
                        if room == "Inside":
                            if price == "SOLD OUT":
                                inside = "N/A"
                            else:
                                inside = price.replace('$', '').replace(',', '')
                        elif room == "Oceanview":
                            if price == "SOLD OUT":
                                oceanview = "N/A"
                            else:
                                oceanview = price.replace('$', '').replace(',', '')
                        elif room == "Balcony":
                            if price == "SOLD OUT":
                                balcony = "N/A"
                            else:
                                balcony = price.replace('$', '').replace(',', '')
                        elif room == "Princess Grill":
                            if price == "SOLD OUT":
                                suite = "N/A"
                            else:
                                suite = price.replace('$', '').replace(',', '')
                    temp = [de, de, vessel_id, vessel_name, cruise_id, cruise_line_name,
                            itinerary_id, brochure_name, duration, sail_date, return_date, inside,
                            oceanview, balcony, suite]
                    if [vessel_name, sail_date, return_date] in names:
                        if inside == 'N/A' and oceanview == 'N/A' and balcony == 'N/A' and suite == 'N/A':
                            continue
                        else:
                            to_write.append(temp)
                            print(temp)
                    else:
                        names.append([vessel_name, sail_date, return_date])
                        to_write.append(temp)
                        print(temp)
                except AttributeError:
                    print("getting alternative pricing on:", vessel_name, sail_date, return_date)
                    button_url = 'http://www.cunard.com' + button_url + '/staterooms-and-fares/'
                    prices_page = session.get(url=button_url).text
                    soup = BeautifulSoup(prices_page, 'lxml')
                    gradelist = soup.find('ul', {'class': 'gradeList gradeListWithBg'})
                    try:
                        lis = gradelist.find_all('li')
                    except AttributeError:
                        print("Skipping cause bad pricing")
                        continue
                    inside = "N/A"
                    oceanview = "N/A"
                    balcony = "N/A"
                    suite = "N/A"
                    for l in lis:
                        navlink = l.find('span')
                        room = navlink.find('strong').text
                        spans = navlink.find_all('span')
                        price = spans.contents[1].text.split('$')[1].replace(' pp', '')
                        if room == "Inside":
                            if price == "SOLD OUT":
                                inside = "N/A"
                            else:
                                inside = price.replace('$', '').replace(',', '')
                        elif room == "Oceanview":
                            if price == "SOLD OUT":
                                oceanview = "N/A"
                            else:
                                oceanview = price.replace('$', '').replace(',', '')
                        elif room == "Balcony":
                            if price == "SOLD OUT":
                                balcony = "N/A"
                            else:
                                balcony = price.replace('$', '').replace(',', '')
                        elif room == "Princess Grill":
                            if price == "SOLD OUT":
                                suite = "N/A"
                            else:
                                suite = price.replace('$', '').replace(',', '')
                    temp = [de, de, vessel_id, vessel_name, cruise_id, cruise_line_name,
                            itinerary_id, brochure_name, duration, sail_date, return_date, inside,
                            oceanview, balcony, suite]
                    if [vessel_name, sail_date, return_date] in names:
                        if inside == 'N/A' and oceanview == 'N/A' and balcony == 'N/A' and suite == 'N/A':
                            continue
                        else:
                            to_write.append(temp)
                            print(temp)
                    else:
                        names.append([vessel_name, sail_date, return_date])
                        to_write.append(temp)
                        print(temp)
        index += 1


def parse_bad(de):
    url = ''
    try:
        url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=1&rt=&sort='
        count = int(session.get(url=url).json()['Count'])
        pages = math.ceil(count / 5)
    except JSONDecodeError:
        try:
            url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=1&rt=&sort='
            count = int(session.get(url=url).json()['Count'])
            pages = math.ceil(count / 5)
        except JSONDecodeError:
            print("skipped major url:", url)
            failed_urls.append(de)
            return
    index = 1
    while index <= int(pages):
        try:
            url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=' + str(
                index) + '&rt=&sort='
            page = session.get(url=url).json()['Html']
        except JSONDecodeError:
            try:
                url = 'http://www.cunard.com/cruise-search/book-a-cruise/results/?tids=' + de + '&pg=' + str(
                    index) + '&rt=&sort='
                page = session.get(url=url).json()['Html']
            except JSONDecodeError:
                print("skipped", url)
                return
        soup = BeautifulSoup(page, 'lxml')
        results_on_page = soup.find_all('li', {'class': 'resultRow'})
        for result in results_on_page:
            brochure_name = result.find('h2').text
            vessel_name = result.find('ul', {'class': 'resultOverview'}).find('img')['alt']
            result_overview = result.find('ul', {'class': 'resultOverview'})
            duration = result_overview.contents[1].text.split()[0]
            dates = result_overview.contents[2].text.split(" - ")
            yr = dates[1].split()[2]
            sail_date = convert_date(dates[0], yr)
            return_date = convert_return(dates[1])
            button_url = ''
            itinerary_id = ''
            cruise_id = ''
            vessel_id = ''
            cruise_line_name = 'Cunard Cruises'
            try:
                buttons = result.find('div', {'class': 'pricing'}).find_all('a')
                if len(buttons) > 1:
                    button_url = buttons[1]['href']
                else:
                    button_url = buttons[0]['href']
            except KeyError:
                print('Failed:', vessel_name, dates)
            if 'http' in button_url:
                inside = "N/A"
                oceanview = "N/A"
                balcony = "N/A"
                suite = "N/A"
                temp = [de, de, vessel_id, vessel_name, cruise_id, cruise_line_name,
                        itinerary_id, brochure_name, duration, sail_date, return_date, inside,
                        oceanview, balcony, suite]
                if [vessel_name, sail_date, return_date] in names:
                    continue
                else:
                    to_write.append(temp)
                    print(temp)
            else:
                button_url = 'http://www.cunard.com' + button_url
                prices_page = session.get(url=button_url).text
                soup = BeautifulSoup(prices_page, 'lxml')
                gradelist = soup.find('ul', {'class': 'gradeList'})
                try:
                    lis = gradelist.find_all('li')
                    inside = "N/A"
                    oceanview = "N/A"
                    balcony = "N/A"
                    suite = "N/A"
                    for l in lis:
                        navlink = l.find('span')
                        room = navlink.find('strong').text
                        price = navlink.find('em').text
                        if room == "Inside":
                            if price == "SOLD OUT":
                                inside = "N/A"
                            else:
                                inside = price.replace('$', '').replace(',', '')
                        elif room == "Oceanview":
                            if price == "SOLD OUT":
                                oceanview = "N/A"
                            else:
                                oceanview = price.replace('$', '').replace(',', '')
                        elif room == "Balcony":
                            if price == "SOLD OUT":
                                balcony = "N/A"
                            else:
                                balcony = price.replace('$', '').replace(',', '')
                        elif room == "Princess Grill":
                            if price == "SOLD OUT":
                                suite = "N/A"
                            else:
                                suite = price.replace('$', '').replace(',', '')
                    temp = [de, de, vessel_id, vessel_name, cruise_id, cruise_line_name,
                            itinerary_id, brochure_name, duration, sail_date, return_date, inside,
                            oceanview, balcony, suite]
                    if [vessel_name, sail_date, return_date] in names:
                        if inside == 'N/A' and oceanview == 'N/A' and balcony == 'N/A' and suite == 'N/A':
                            continue
                        else:
                            to_write.append(temp)
                            print(temp)
                    else:
                        names.append([vessel_name, sail_date, return_date])
                        to_write.append(temp)
                        print(temp)
                except AttributeError:
                    print("getting alternative pricing on:", vessel_name, sail_date, return_date)
                    button_url = 'http://www.cunard.com' + button_url + '/staterooms-and-fares/'
                    prices_page = session.get(url=button_url).text
                    soup = BeautifulSoup(prices_page, 'lxml')
                    gradelist = soup.find('ul', {'class': 'gradeList gradeListWithBg'})
                    lis = gradelist.find_all('li')
                    inside = "N/A"
                    oceanview = "N/A"
                    balcony = "N/A"
                    suite = "N/A"
                    for l in lis:
                        navlink = l.find('span')
                        room = navlink.find('strong').text
                        spans = navlink.find_all('span')
                        price = spans.contents[1].text.split('$')[1].replace(' pp', '')
                        if room == "Inside":
                            if price == "SOLD OUT":
                                inside = "N/A"
                            else:
                                inside = price.replace('$', '').replace(',', '')
                        elif room == "Oceanview":
                            if price == "SOLD OUT":
                                oceanview = "N/A"
                            else:
                                oceanview = price.replace('$', '').replace(',', '')
                        elif room == "Balcony":
                            if price == "SOLD OUT":
                                balcony = "N/A"
                            else:
                                balcony = price.replace('$', '').replace(',', '')
                        elif room == "Princess Grill":
                            if price == "SOLD OUT":
                                suite = "N/A"
                            else:
                                suite = price.replace('$', '').replace(',', '')
                    if [vessel_name, sail_date, return_date] in names:
                        if inside == 'N/A' and oceanview == 'N/A' and balcony == 'N/A' and suite == 'N/A':
                            continue
                        else:
                            to_write.append(temp)
                            print(temp)
                    else:
                        names.append([vessel_name, sail_date, return_date])
                        to_write.append(temp)
                        print(temp)
        index += 1


pool.map(parse, destinations)
pool.close()
pool.join()

if len(major_urls) > 0:
    print("Retrying failed (" + str(len(major_urls)) + '')
    pool2.map(parse_bad, major_urls)
    pool2.close()
    pool2.join()
if len(failed_urls) > 0:
    print("Re-Retrying failed (" + str(len(failed_urls)) + '')
    pool3.map(parse_bad, failed_urls)
    pool3.close()
    pool3.join()


def write_file(data_array):
    userhome = os.path.expanduser('~')
    print(userhome)
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Cunard Cruises.xlsx'
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
                    try:
                        date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, money_format)
                    except ValueError:
                        split = str(en).split('/')
                        en = split[1] + '/' + split[0] + '/' + split[2]
                        date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except TypeError:
                    worksheet.write_string(row_count, column_count, en, centered)
            if column_count == 10:
                try:
                    try:
                        date_time = datetime.datetime.strptime(str(en), "%m/%d/%Y")
                        worksheet.write_datetime(row_count, column_count, date_time, money_format)
                    except ValueError:
                        split = str(en).split('/')
                        en = split[1] + '/' + split[0] + '/' + split[2]
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


write_file(to_write)
input("Press any key to continue...")
