import datetime
import os
import time
from multiprocessing.dummy import Pool as ThreadPool

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, NoSuchElementException

tested = []


# def get_proxy():
#     print("Looking for a working proxy server")
#     soup = BeautifulSoup(requests.get('https://www.us-proxy.org').text, 'lxml')
#     table = soup.find('table', {'id': 'proxylisttable'})
#     tbody = table.find('tbody')
#     proxies = []
#     for tr in tbody:
#         columns = tr.find_all('td')
#         if columns[2].text in 'US' and columns[4].text in 'elite proxy' and columns[6].text in 'yes':
#             proxies.append("https://" + columns[0].text + ":" + columns[1].text)
#     for p in proxies:
#         try:
#             proxy_line = {'https': p}
#             resp = requests.post('https://www.princess.com', proxies=proxy_line,
#                                  timeout=10)
#             if resp.ok:
#                 if proxy_line not in tested:
#                     print("Found one!")
#                     return proxy_line
#                 else:
#                     print("Found one banned")
#             else:
#                 print(p, "Not working")
#         except requests.exceptions.ProxyError:
#             print(p, "Not working")
#         except requests.exceptions.ConnectTimeout:
#             print(p, "Not working")
#         except requests.exceptions.ReadTimeout:
#             print(p, "Not working")


session = requests.session()
test_url = 'https://www.princess.com/find/json/getJsonProducts.do'
# profile = webdriver.FirefoxProfile()
# proxy = get_proxy()
# for_driver = proxy['https'].replace('https://', '').split(':')
# profile.set_preference("network.proxy.type", 1)
# profile.set_preference("network.proxy.ssl", for_driver[0])
# profile.set_preference("network.proxy.ssl_port", int(for_driver[1]))
# profile.update_preferences()
h = session.get(test_url)
cookie = session.cookies.get_dict()
ports_set = set()

cruises = []
pool = ThreadPool(2)
body = 'searchCriteria.subTrades%5B0%5D=&searchCriteria.sortBy=L&searchCriteria.versionB=false&searchCriteria.startDateRange=&searchCriteria.endDateRange=&searchCriteria.searchKey=bb0b9ce2-7ccc-411c-a33c-b4fba3764566&searchCriteria.meta=I&searchCriteria.pastPax=false&searchCriteria.noOfPax=2&searchCriteria.cruiseTour=false&searchCriteria.itineraryCode=&searchCriteria.voyageCode=&searchCriteria.tourCode=&searchCriteria.startIndex=0&searchCriteria.endIndex=440&searchCriteria.positionIndex=0&pageName=searchresult&ubeData.ubeId=PCDIR&searchCriteria.currency=USD&searchCriteria.countryForPrice=BG&searchCriteria.cruiseDetail=false&searchCriteria.shipVersion=&searchCriteria.webDisplay=Y&searchCriteria.applyCoupon=true&searchCriteria.ocean=&searchCriteria.newVersion=false'
headers = {
    'Host': 'www.princess.com',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
    'Accept': '*/*',
    'Accept-Language': 'bg-BG,bg;q=0.8,ru;q=0.6,en;q=0.4',
    'Accept-Encoding': 'gzip, deflate',
    "Referer": "https://www.princess.com/find/searchResults.do",
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'X-Requested-With': 'XMLHttpRequest',
    'Content-Length': '776',
    'Cookie': 'offloadcookie=!uaF821aYSDstaSSFJhHqn5/ouTmcBVYGbN75zKEV64/WsW8ukBHrsFqaBsURMfPnl4wBKQKdPSQ/Bw==; ak_bmsc=D9EF391211F2784EB363FBA4E083675E02142D61B7460000A91AB55A8A392F5B~plyphhtp8DwuXQgB+DEBtGhv7e9R9rMF7upkSBzH7OlYTME9avJfeehf3f9PKBUyfMypJ4T6VUi/oVcY567qofur/PPpZ161VTYcCWTIgyfpNxiyk4l//MaBgTXLoVMBLjPT3dwKbc7+VC7X9DB29bQxUA1tCwAoOTM8rdQ3lwcn/ReBspkjiDZjsukSd1pz7Jxkm2jCGN+xeonggDXTo3fLA/qSe17a1pC5C9g9m5laM=; bm_sz=0CAF96DB723C308E31AEC303B54E7BE9~QAAQYS0UAq16qVBiAQAAzSVwU7mahrap43C3ovqC0nBHMHy6bD4XVnxWb+atnsZfV7QnFpDgnpJNHOad2JxuAEU4afwbpCbAqvVB0AUQkLT0hQQAh9XyWn6amK+uBu2fraz6lZ3KVlrO9M1JEOGsLdQLZaXwyxr/V5vrwOdxxQlXFdPaU40EMtoQ7MJ768R11w==; COOKIE_CHECK=YES; check=true; AMCVS_21C91F2F575539D07F000101%40AdobeOrg=1; _fipc_=BG; _fipa_=SOF; loc=SH3HZ63UPZQNUDCIAMO6Y525HJ5LLE7BPS3W3YETQZTFIV2CTVGYFFRJU4VDKCZKMNJA5REBAXXRPIUXPXIXDUAD6Q; getLocale=%7B%22lastUpdated%22%3A%221521818288%22%2C%22ipAddress%22%3A%2294.26.58.247%22%2C%22zipCode%22%3A%22%22%2C%22brochures%22%3A%22false%22%2C%22specialOffers%22%3A%22false%22%2C%22status%22%3A%22%22%2C%22primaryCurrency%22%3A%22USD%22%2C%22aircity%22%3A%22SOF%22%2C%22country%22%3A%22BG%22%2C%22countryPhone%22%3A%22%22%2C%22isEU%22%3A%22true%22%7D; s_dfa=crbrprincessprodus%2Ccrbrcarnivalbrandsprodus; _abck=43DDF4FB6EB06E157D091B61F1125EEF02142D61B7460000A91AB55A20618423~0~eW8Ij1qyDNFxgCUuN8fp5Qq7sgzVQ/V8F+aVhk8dYLk=~-1~-1; _ga=GA1.2.1229489472.1521818289; _gid=GA1.2.782976354.1521818289; _dc_gtm_UA-4086206-54=1; AMCV_21C91F2F575539D07F000101%40AdobeOrg=1406116232%7CMCIDTS%7C17614%7CMCMID%7C38318429095190454353487143113038455123%7CMCAAMLH-1522423088%7C6%7CMCAAMB-1522423088%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1521825488s%7CNONE%7CMCAID%7CNONE%7CMCSYNCSOP%7C411-17621%7CvVersion%7C2.5.0; JSESSIONID=0000Lu4LHZHmsPOme4eRhdQ8ztu:13ehlrk2j; IBMID=Lu4LHZHmsPOme4eRhdQ8ztu:1; _gat_UA-4086206-54=1; persistValue=null; gds_s=First%20Visit; s_vnum=1522530000719%26vn%3D1; s_invisit=true; s_cc=true; aam_analytics=segment%3D5554399; at_carnivalbrands=segments%3D5554399; edge_check=anyone%3Dtrue%2Cvisitor%3Dcheck; aam_uuid=38589892945040043833514004724511820151; s_sq=%5B%5BB%5D%5D; ADRUM=s=1521818303750&r=https%3A%2F%2Fwww.princess.com%2F%3F0; akamai-pparams=noOfPax=2; mbox=session#2270be26f13d4217b80bc243cab450e2#1521820167|PC#2270be26f13d4217b80bc243cab450e2.26_16#1585063090; dtmEntryPage=pcl:cruise_shopping:search_results; s_ppn=pcl%3Acruise_shopping%3Asearch_results; _fby_site_=1%7Cprincess.com%7C1521818289%7C1521818289%7C1521818289%7C1521818308%7C1%7C2%7C2; __atuvc=2%7C12; __atuvs=5ab51ab0421884b8001; s_nr=1521818313157-New; gds=1521818313158',
    'Connection': 'keep-alive',
    'Origin': 'https://www.princess.com',
    'ADRUM': 'isAjax:true'

}
test_headers = {
    'Host': 'www.princess.com',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0',
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.5',
    'Accept-Encoding': 'gzip, deflate',
    "Referer": "http://www.princess.com/find/searchResults.do",
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'X-Requested-With': 'XMLHttpRequest',
    'Content-Length': '774',
    'Cookie': '_aeu=QCQ=; _aes=QSE=; dl.VoyageCode=0:S709C; getLocale=%7B%22specialOffers%22%3A%22false%22%2C%22status%22%3A%22%22%2C%22primaryCurrency%22%3A%22USD%22%2C%22country%22%3A%22BG%22%2C%22countryPhone%22%3A%22%22%2C%22isEU%22%3A%22true%22%2C%22brochures%22%3A%22false%22%2C%22lastUpdated%22%3A1487698256566%7D; _aeu=QCQ=; uk_ok=true; JSESSIONID=0001yy0unZLr4F7z2WNNdpq04hH:181iit3or; _fipc_=US; loc=SH3HZ63UPZQNU3YJRB5MKYKKGV5LLE7BPS3W3YEN6FPNOYMSVVAFFLE5AVVLOCR4XQY4G7YE3WZ2PSD2QO44QUUSME; ak_bmsc=518181399E7438AC1DC1428EFB0426A55C7AD525931E000071EDAC582D97CE49~plXog1ZCtAlVzF6i1DIivuXk+Qdj1iYQTJ+fVG+e9vmqPwTO8TI5uQW9Pz0G49NTPeYLrz4c2lBszS5P8x5lVU9SSqWrI85jN9VR0DakaIDKztTXo6/hGhMuOxYcMOSqSVOuL4xEH5atAmO8SDtEFNLOI8VhAHZQ7yCNTdB6PbJyGlC148ieos0ccMegH81Ct3PGHk7JKA+u9dUFa4hZbH9BkoGmq31ejPyWsnDud9jCg=; COOKIE_CHECK=YES; __utmt_princess=1; __utmt=1; _dc_gtm_UA-4086206-54=1; EG-S-ID=A413e4a93c-321c-4fc6-8e05-5275c4ceed2e; EG-U-ID=A7f20e6cbf-6769-4cc3-a6e3-2574c50824b9; spo=2QRFD3E7SBIWZUMGHAH7WTU6YU; _gat_UA-4086206-54=1; AMCVS_21C91F2F575539D07F000101%40AdobeOrg=1; AMCV_21C91F2F575539D07F000101%40AdobeOrg=2121618341%7CMCIDTS%7C17219%7CMCMID%7C45019628955626001270446066839249503603%7CMCAAMLH-1488303056%7C6%7CMCAAMB-1488332791%7CNRX38WO0n5BH8Th-nqAG_A%7CMCOPTOUT-1487735191s%7CNONE%7CMCAID%7CNONE; s_vnum=1488319200028%26vn%3D3; s_sq=%5B%5BB%5D%5D; visitNumTemp=3; numberofnights=2; targetVoyageIDTemp=S709C; rvc=JUD2DUEVN4HWOZYNRYWG7QPJZFFRQTKNVZHJQAVBDCTZ7ICBQVKQ; booking_engine_used=PCDIR; search_counter=2; _ga=GA1.2.1000495142.1487698257; _fby_site_=1%7Cprincess.com%7C1487727989%7C1487727989%7C1487727989%7C1487728182%7C1%7C3%7C3; edge_check=visitor%3Dprincess%2Canyone%3Dtrue%2Cvisitor%3Dcheck; at_carnivalbrands=segments%3D5554399; aam_uuid=45291411661163250190472928465419857239; mbox=PC#1d8c575a12c940008cb69e668a969ade.26_18#1488937785|session#5b9f9d2e459e40939028d9b37098e384#1487730045|check#true#1487728245; persistValue=null; mf_f3a02463-b43f-48da-9dcb-90e7d2f103b1=-1; __utma=169354720.1000495142.1487698257.1487703940.1487727986.3; __utmb=169354720.13.9.1487728184815; __utmc=169354720; __utmz=169354720.1487698257.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); s_dfa=crbrprincessprodus%2Ccrbrcarnivalbrandsprodus; s_cc=true; __atuvc=5%7C8; __atuvs=58aced74c6d79c7b002; s_ppvl=pcl%253Amy_princess%253Aspecialofferregistration%253A%2C0%2C0%2C1%2C1%2C1%2C1920%2C1080%2C1%2CP; s_ppv=pcl%253Acruise_shopping%253Asearch_results%2C85%2C50%2C16477%2C1001%2C948%2C1920%2C1080%2C1%2CP; s_ppn=pcl%3Acruise_shopping%3Asearch_results; s_nr=1487728278191-Repeat; gds=1487728278192; gds_s=Less%20than%201%20day; s_invisit=true',
    'Connection': 'keep-alive',
    'Pragma': 'no-cache',
    'Cache-Control': 'no-cache'
}


def fix_header():
    browser = webdriver.Firefox(executable_path='geckodriver.exe')
    browser.get('https://www.princess.com/find/searchResults.do')
    try:
        browser.find_element_by_xpath(
                    "//*[@class='expand-table view-all-link util-link plain-text-btn gotham-bold']").click()
    except ElementClickInterceptedException:
        try:
            browser.find_element_by_xpath("//*[@class='action-toolbar-btn']").click()
            browser.find_element_by_xpath(
                "//*[@class='expand-table view-all-link util-link plain-text-btn gotham-bold']").click()
        except ElementClickInterceptedException:
            try:
                browser.find_element_by_xpath("//img[@src='/images/global/lightbox_close-btn.png']").click()
                time.sleep(5)
                browser.find_element_by_xpath("//*[@class='action-toolbar-btn']").click()
                browser.find_element_by_xpath(
                    "//*[@class='expand-table view-all-link util-link plain-text-btn gotham-bold']").click()
            except Exception:
                try:
                    browser.find_element_by_class_name('mfp-close').click()
                    browser.find_element_by_xpath("//img[@src='/images/global/lightbox_close-btn.png']").click()
                    time.sleep(5)
                    browser.find_element_by_xpath("//*[@class='action-toolbar-btn']").click()
                    browser.find_element_by_xpath(
                        "//*[@class='expand-table view-all-link util-link plain-text-btn gotham-bold']").click()
                except NoSuchElementException:

                    browser.find_element_by_xpath("//*[@class='action-toolbar-btn']").click()
                    time.sleep(2)
                    browser.find_element_by_xpath(
                        "//*[@class='expand-table view-all-link util-link plain-text-btn gotham-bold']").click()
    cookies = browser.get_cookies()
    browser.close()
    chrome_cookie = ''
    for c in cookies:
        chrome_cookie += c['name']
        chrome_cookie += '='
        chrome_cookie += c['value']
        chrome_cookie += "; "
    return chrome_cookie[:-2]


headers['Cookie'] = fix_header()
url = "https://www.princess.com/find/json/getJsonProducts.do?"
page = session.get(url)
missed = []
cruise_data = page.json()
for key, value in cruise_data['data'].items():
    cruises.append([key, value['I'], value['M']])
url = 'https://www.princess.com/find/pagination.do'
page = session.post(url, data=body, headers=headers)
soup = BeautifulSoup(page.text, 'lxml')
price_url = ''
itin_id = ''
to_write = []
codes = []
itins = soup.find_all('div', {'class': 'result'})


def preformat_date(unformated):
    splitter = unformated.split(', ')
    exact = splitter[1].split()
    day = exact[1]
    if day[0] == '0':
        day = day[1]
    month = exact[0]
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
    calculated = calculated.strftime("%m/%d/%Y")
    return calculated


def parse(i):
    itin_id = i['id']
    cruise_days = i.find('div', {'class': 'cruise-days'})
    days = cruise_days.find('div').text
    dest = ''
    for c in cruises:
        if c[0] == itin_id:
            dest = c[2]
            cruises.remove(c)
    ports_block = i.find('div', {'class': 'row ports-info'})
    raw_ports = ports_block.find_all('a')
    ports = []
    for port in raw_ports:
        ports.append(port.text.strip())
        ports_set.add(port.text.strip())
    title = i.find('a')
    brochure_name = title.text
    url = 'https://www.princess.com/find/viewAllCruises.do'
    data = 'searchCriteria.subTrades%5B0%5D=&searchCriteria.sortBy=L&searchCriteria.versionB=false&searchCriteria.startDateRange=&searchCriteria.endDateRange=&searchCriteria.searchKey=b182db8c-1bbc-4cdf-92a7-be23ce97b87b&searchCriteria.meta=I&searchCriteria.pastPax=false&searchCriteria.noOfPax=2&searchCriteria.cruiseTour=false&searchCriteria.itineraryCode=' + itin_id + '&searchCriteria.voyageCode=&searchCriteria.tourCode=&searchCriteria.startIndex=30&searchCriteria.endIndex=40&searchCriteria.positionIndex=0&pageName=searchresult&ubeData.ubeId=PCDIR&searchCriteria.currency=USD&searchCriteria.countryForPrice=US&searchCriteria.cruiseDetail=false&searchCriteria.shipVersion=&searchCriteria.webDisplay=Y&searchCriteria.applyCoupon=true&searchCriteria.ocean=&searchCriteria.newVersion=false'
    prices_page = session.post(url, data=data, headers=headers)
    info = BeautifulSoup(prices_page.text, 'lxml')
    sailings = []
    tables = info.find_all('table', {'class': 'pricing-table'})
    for table in tables:
        rows = table.find_all('tr')
        for row in rows:
            tds = row.find_all('td')
            if len(tds) == 5:
                divs = row.find_all('div', limit=6)
                temp = []
                for d in divs:
                    temp.append(d.text.strip())
                container = table.parent.parent
                button = container.find('button', {'value': 'Cruise Details'})['id']
                # button = table.find('button', {'value': 'Cruise Details'})['id']
                temp.append(button.split('-')[3])
                if len(temp) == 7:
                    sailings.append(temp)
    for sail in sailings:
        # try:
        sail_date = preformat_date(sail[0])
        return_date = calculate_days(sail_date, days)
        inside = sail[2].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
        oceanview = sail[3].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
        balcony = sail[4].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
        suite = sail[5].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
        code = sail[6]
        destination_code = dest
        destination_name = dest
        vessel_id = ''
        vessel_name = sail[1]
        cruise_line_name = "Princess Cruises"
        cruise_id = ''
        number_of_nights = days
        itin_id = ''
        if 'Cross International Date Line' in ports:
            detail_url = 'https://www.princess.com/find/displayItineraryDetails.do'
            details_body = 'searchCriteria.voyageCode=' + code + ''
            detail_page = session.post(url=detail_url, headers=headers, data=details_body).text
            detail_soup = BeautifulSoup(detail_page, 'lxml')
            table = detail_soup.find('table', {'class': 'ports-table'})
            rows = table.find_all('tr')
            span = rows[len(rows) - 1]
            actual = span.find_all('td')[0].text.strip().split()[2]
            itin_id = str((int(actual) - int(return_date.split("/")[1])))
            split = return_date.split('/')
            return_date = split[0] + '/' + actual + '/' + split[2]
        ports_string = ''
        for single in ports:
            ports_string += ", " + single
        temp = [destination_code, destination_name, vessel_id, vessel_name, cruise_id, cruise_line_name,
                itin_id, brochure_name, number_of_nights, sail_date, return_date, inside,
                oceanview, balcony, suite, ports_string[1:]]
        to_write.append(temp)
        print(temp)
        # except IndexError:
        #     print(data)


pool.map(parse, itins)
pool.close()
pool.join()
print('Excepions', len(cruises))
for cr in cruises:
    cruise_line_name = "Princess Cruises"
    vessel_id = ''
    url = 'https://www.princess.com/find/viewAllCruises.do'
    data = 'searchCriteria.subTrades%5B0%5D=&searchCriteria.sortBy=L&searchCriteria.versionB=false&searchCriteria.startDateRange=&searchCriteria.endDateRange=&searchCriteria.searchKey=b182db8c-1bbc-4cdf-92a7-be23ce97b87b&searchCriteria.meta=I&searchCriteria.pastPax=false&searchCriteria.noOfPax=2&searchCriteria.cruiseTour=false&searchCriteria.itineraryCode=' + \
           cr[
               0] + '&searchCriteria.voyageCode=&searchCriteria.tourCode=&searchCriteria.startIndex=30&searchCriteria.endIndex=40&searchCriteria.positionIndex=0&pageName=searchresult&ubeData.ubeId=PCDIR&searchCriteria.currency=USD&searchCriteria.countryForPrice=US&searchCriteria.cruiseDetail=false&searchCriteria.shipVersion=&searchCriteria.webDisplay=Y&searchCriteria.applyCoupon=true&searchCriteria.ocean=&searchCriteria.newVersion=false'
    prices_page = session.post(url, data=data, headers=headers)
    info = BeautifulSoup(prices_page.text, 'lxml')
    try:
        sail_date = info.find('div', {'class': 'depart-date gotham-bold no-wrap'}).text
    except AttributeError:
        print("skipping")
    vessel_name = info.find('div', {'class': 'voyage-ship'}).text
    sailings = []
    tables = info.find_all('table', {'class': 'pricing-table'})
    for table in tables:
        rows = table.find_all('tr')
        for row in rows:
            tds = row.find_all('td')
            if len(tds) == 5:
                divs = row.find_all('div', limit=6)
                temp = []
                for d in divs:
                    temp.append(d.text.strip())
                container = table.parent.parent
                button = container.find('button', {'value': 'Cruise Details'})['id']
                # button = table.find('button', {'value': 'Cruise Details'})['id']
                temp.append(button.split('-')[3])
                if len(temp) == 7:
                    sailings.append(temp)
    for sail in sailings:
        try:
            inside = sail[2].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
            oceanview = sail[3].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
            balcony = sail[4].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
            suite = sail[5].split('.')[0].replace('$', '').replace(',', '').replace('Sold Out', 'N/A')
            detail_url = 'https://www.princess.com/find/displayItineraryDetails.do'
            details_body = 'searchCriteria.voyageCode=' + sail[6] + ''
            detail_page = session.post(url=detail_url, headers=headers, data=details_body).text
            detail_soup = BeautifulSoup(detail_page, 'lxml')
            print(detail_soup)
            number_of_nights = detail_soup.find('span').text.split(' | ')[0].split()[0]
            return_date = calculate_days(preformat_date(sail_date), number_of_nights)
            ports = []
            table = detail_soup.find('table', {'class': 'ports-table'})
            rows = table.find_all('tr')
            print(len(rows))
            for row in rows:
                ports.append(row.find_all('td')[1].contents[0].text.strip().split(',')[0])
            if 'Cross International Date Line' in ports:
                span = rows[len(rows) - 1]
                actual = span.find_all('td')[0].text.strip().split()[2]
                itin_id = str((int(actual) - int(return_date.split("/")[1])))
                split = return_date.split('/')
                return_date = split[0] + '/' + actual + '/' + split[2]
            ports_string = ''
            for single in ports:
                ports_string += ", " + single
            temp = [cr[2], cr[2], vessel_id, vessel_name, '', cruise_line_name,
                    itin_id, '', number_of_nights, preformat_date(sail_date), return_date, inside,
                    oceanview, balcony, suite, ports_string[1:]]
            to_write.append(temp)
            print(temp)
        except IndexError:
            print(data)


def write_file_to_excel(data_array):
    userhome = os.path.expanduser('~')
    now = datetime.datetime.now()
    path_to_file = userhome + '\\Dropbox\\XLSX\\' + str(now.year) + '-' + str(now.month) + '-' + str(
        now.day) + '/' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '- Princess Cruises.xlsx'
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
                    date_time = datetime.datetime.strptime(str('9/9/2090'), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
                except ValueError:
                    date_time = datetime.datetime.strptime(str('9/9/2090'), "%m/%d/%Y")
                    worksheet.write_datetime(row_count, column_count, date_time, money_format)
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


write_file_to_excel(to_write)
f = open("ports.txt", 'w')
for row in list(ports_set):
    f.write(row + '\n')
f.close()
input("Press any key to continue...")
