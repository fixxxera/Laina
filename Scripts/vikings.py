from multiprocessing.dummy import Pool as ThreadPool

import requests

session = requests.session()
pool = ThreadPool(10)
body = {
    "Host": "www.vikingcruises.com",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "en-US,en;q=0.5",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": 'https://www.vikingcruises.com/oceans/search-cruises/index.html?Regions=Scandinavia%20%26%20Northern'
               '%20Europe|The%20Americas%20%26%20Caribbean|Mediterranean|Quiet%20Season%20Mediterranean|Asia%20%26'
               '%20Australia|Africa|World%20%26%20Grand%20Voyages',
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Content-Length": "28637",
    "Cookie": "VRCPERM=%7b%22ts%22%3a0%2c%22ss%22%3a0%2c%22ns%22%3a1%2c%22nso%22%3a0%2c%22sv%22%3a0%2c%22svo%22%3a0"
              "%2c%22nsd%22%3a%221901-3-1%22%2c%22nsod%22%3a%221901-3-1%22%2c%22svd%22%3a%222018-5-7%22%2c%22svod%22"
              "%3a%221901-3-1%22%2c%22rvo%22%3a0%2c%22cnl%22%3a0%2c%22cnld%22%3a%222018-5-12%22%2c%22pi%22%3a"
              "%22636612823624286000vcefjqepfai%22%2c%22mv%22%3a%22%22%2c%22sl%22%3a%22%22%2c%22vnd%22%3a%222018-5-7"
              "%22%2c%22ftvac%22%3afalse%2c%22pkey%22%3a%22%22%7d; "
              "utag_main=v_id:01633b72ac7600017792037a86c70104e002400d00d60$_sn:3$_ss:0$_st:1526132496927$dc_visit"
              ":3$ses_id:1526129561308%3Bexp-session$_pn:4%3Bexp-session$_prevpage:voc%3Aus%3Ahome"
              "%3Afind_your_cruise_search%3Bexp-1526134296930$dc_event:14%3Bexp-session; "
              "s_fid=1EDD32F764C485DA-1E773EB1F9A7B76B; s_vnum=1527800400464%26vn%3D3; "
              "_bcvm_vrid_1399666957461524825"
              "=834038834948252468T829DCBB6CE39F01E4B882D95E37E1B0AB7BF0E23E27136031F36560B8E4E85D2651C36524D1801859147FD97944BA6696A378B4D2A464F2DA76ECC7871025B64; CountryCode=BG; BYPASS=true; bypassbrowsercompatibilitycheck=true; VRCSESS=%7b%22pd%22%3a4%2c%22ga%22%3a0%2c%22go%22%3a0%2c%22ps%22%3a0%2c%22rv%22%3a0%2c%22svt%22%3a0%2c%22nst%22%3a0%2c%22hp%22%3a0%2c%22svo%22%3a0%2c%22nso%22%3a0%2c%22pu%22%3a0%2c%22si%22%3a%22%22%2c%22fv%22%3a%22%22%2c%22ui%22%3a%22571692000rxhtar%22%2c%22bid%22%3a%22%22%2c%22vpp%22%3a%22%22%2c%22sfym%22%3a%22%22%2c%22sfoff%22%3a%22%22%2c%22sftcm%22%3a%22%22%7d; bypasscookiedisclaimer=true; s_cc=true; s_invisit=true; s_visit=1; t_camppath=typed-unknown; t_chnlpath=typed-unknown; s_sq=%5B%5BB%5D%5D; _bcvm_vid_1399666957461524825=834043023801157423T4A7786EA8E9C5BCF96C2931901B73ED698DB2437F3E6A6DF35F4DD4CBBF7430686E33E48FF3D812FFDB265F47D6F660C0CCCE5986FB725B4E3779BF9551CE93A; bc_pv_end=",
    "Connection": "keep-alive"
}
response = session.post('https://www.vikingcruises.com/oceans/fypc/GetSuperFacAsync', data=body)
to_write = []
root_voyages = []
for voyage in response.json()['Value']['Results']:
    title = voyage['VoyageName']
    duration = voyage['Days']
    ports = voyage['ItineraryDays']
    destination = voyage['Regions']
    url = voyage['PageUrl']
    id = voyage['TcmId']
    root_voyages.append([id, title, duration, ports, destination, url])


def parse(voyage):
    request_body = {"cruiseId": voyage[0], "parameters": ""}
    response = session.post('https://www.vikingcruises.com/oceans/ECommerce/DnPCruiseFullInfoV3?v=11',
                            data=request_body)
    for sailing in response.json()['cruises']:
        request_body = {"cruiseId": voyage[0], "parameters": "", "sailingKey": sailing['sailingKey']}
        detail_response = session.post('https://www.vikingcruises.com/oceans/ECommerce/DnPSailingDetailsV3?v=11',
                                       data=request_body).json()
        ship = detail_response['sailingData']['cruiseShip']['name']
        prices = detail_response['sailingData']['cruiseSuites']


pool.map(parse, root_voyages)
pool.close()
pool.join()