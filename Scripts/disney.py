import requests

headers = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:52.0) Gecko/20100101 Firefox/52.0",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Host": "www.hollandamerica.com",
    "Adrum": "isAjax:true",
    "Refer": "https://disneycruise.disney.go.com/cruises-destinations/list/",
    "Authorization": "BEARER jRulZsQ8iRMkU-ekrfk5-w",
    "Cookie": "localeCookie_jar=%7B%22contentLocale%22%3A%22en%22%2C%22version%22%3A%222%22%2C%22precedence%22%3A0%2C%22localeCurrency%22%3A%22USD%22%2C%22preferredRegion%22%3A%22en%22%7D; party_mix=%5B%7B%22isDefault%22%3Atrue%2C%22accessible%22%3Afalse%2C%22adultCount%22%3A2%2C%22childCount%22%3A0%2C%22seniorCount%22%3A0%2C%22nonAdultAges%22%3A%5B%5D%2C%22orderBuilderId%22%3Anull%2C%22partyMixId%22%3A%220%22%7D%5D; Conversation_UUID=2ce14c80-1eeb-11e7-a41e-79d21e033a5b; geoipLegacy=eyJhcmVhY29kZSI6IjUwNCIsImNvdW50cnkiOiJ1bml0ZWQgc3RhdGVzIiwiY29udGluZW50IjoibmEiLCJjb25uZWN0aW9uIjoiY2FibGUiLCJjb3VudHJ5Y29kZSI6Ijg0MCIsImNvdW50cnlpc29jb2RlIjoidXNhIiwiZG9tYWluIjoiY294Lm5ldCIsImRzdCI6InkiLCJpc3AiOiJjb3ggY29tbXVuaWNhdGlvbnMgaW5jLiIsIm1ldHJvIjoibmV3IG9ybGVhbnMiLCJtZXRyb2NvZGUiOiI2MjIiLCJvZmZzZXQiOiItNTAwIiwicG9zdGNvZGUiOiI3MDExOSIsInNpYyI6ImludGVybmV0IHNlcnZpY2UiLCJzaWNjb2RlIjoiNzM3NDE1Iiwic3RhdGUiOiJsb3Vpc2lhbmEiLCJ6aXAiOiI3MDExOSIsImlwIjoiMjQuMjUyLjEyNS4xNzgifQ%3D%3D; geoip=YToxODp7czo4OiJhcmVhY29kZSI7czozOiI1MDQiO3M6NzoiY291bnRyeSI7czoxMzoidW5pdGVkIHN0YXRlcyI7czo5OiJjb250aW5lbnQiO3M6MjoibmEiO3M6MTA6ImNvbm5lY3Rpb24iO3M6NToiY2FibGUiO3M6MTE6ImNvdW50cnljb2RlIjtzOjM6Ijg0MCI7czoxNDoiY291bnRyeWlzb2NvZGUiO3M6MzoidXNhIjtzOjY6ImRvbWFpbiI7czo3OiJjb3gubmV0IjtzOjM6ImRzdCI7czoxOiJ5IjtzOjM6ImlzcCI7czoyMzoiY294IGNvbW11bmljYXRpb25zIGluYy4iO3M6NToibWV0cm8iO3M6MTE6Im5ldyBvcmxlYW5zIjtzOjk6Im1ldHJvY29kZSI7czozOiI2MjIiO3M6Njoib2Zmc2V0IjtzOjQ6Ii01MDAiO3M6ODoicG9zdGNvZGUiO3M6NToiNzAxMTkiO3M6Mzoic2ljIjtzOjE2OiJpbnRlcm5ldCBzZXJ2aWNlIjtzOjc6InNpY2NvZGUiO3M6NjoiNzM3NDE1IjtzOjU6InN0YXRlIjtzOjk6ImxvdWlzaWFuYSI7czozOiJ6aXAiO3M6NToiNzAxMTkiO3M6MjoiaXAiO3M6MTQ6IjI0LjI1Mi4xMjUuMTc4Ijt9Ow%3D%3D; siteId=dcl; BIGipServer~WDPRO~pool-WDPRO_VARNISHCACHE-PRODB=327665162.20480.0000; PHPSESSID=gnaj8difbi5imu233snuq40qb2; AFFILIATIONS_jar=%7B%22dcl%22%3A%7B%22storedAffiliations%22%3A%5B%5D%7D%7D; bkSent=true; connect.sid=s%3AQLlwd1YcbrLjRFFa-AtIeCSOyKBobN2b.2Mo0NTSQZRZ9hxcoJs%2BIntevUxZkoPYbqwFVlUOiHRY; UNID=6ede07ae-5b37-4bc6-bb54-2e3903707142; UNID=6ede07ae-5b37-4bc6-bb54-2e3903707142; _sdsat_enableClickTale=true; __CT_Data=gpv=4&apv_32300_www07=4; WRUIDB20170404=0; surveyThreshold_jar=%7B%22pageViewThreshold%22%3A4%7D; pep_oauth_token=jRulZsQ8iRMkU-ekrfk5-w; pep_oauth_token_expiration=351; cartIdMapping=%7B%22new%22%3A%227e1d89db-cab5-40de-9429-f1fb85b567f5%22%7D; s_pers=%20s_fid%3D2EDF6E69BAE8BEE4-00462FB5638A2F49%7C1555010800308%3B%20s_gpv_pn%3Dwdpro%252Fdcl%252Fus%252Fen%252Fcommerce%252Fbooking%252Fconsumer%252Fsearchresults%7C1491940600320%3B; s_sess=%20prevPageLoadTime%3Dundefined%257C10.0%3B%20s_cc%3Dtrue%3B%20s_ppv%3D-%3B%20s_wdpro_lid%3D%3B%20s_sq%3D%3B; s_vi=[CS]v1|2C7490350507CFCD-4000010400002121[CE]; mbox=PC#1491673189076-745650.17_48#1499714816|session#1491938111700-270187#1491940676|mboxEdgeServer#mboxedge17.tt.omtrdc.net#1491940676|check#true#1491938876; boomr_rt=""; WDPROView=%7B%22version%22%3A2%2C%22preferred%22%3A%7B%22device%22%3A%22desktop%22%2C%22screenWidth%22%3A1280%2C%22screenHeight%22%3A1024%7D%2C%22deviceInfo%22%3A%7B%22device%22%3A%22desktop%22%2C%22screenWidth%22%3A1280%2C%22screenHeight%22%3A1024%7D%2C%22browserInfo%22%3A%7B%22agent%22%3A%22Chrome%22%2C%22version%22%3A%2257.0.2987.133%22%7D%7D; ADRUM_BTa=R:61|g:75564e8b-df9e-4a56-9d26-ba88842f4fa4; ADRUM_BT1=R:61|i:212|e:21"
}
json = {"currency": "USD", "affiliations": [], "partyMix": [
    {"accessible": False, "adultCount": 2, "seniorCount": 0, "childCount": 0, "orderBuilderId": None,
     "nonAdultAges": [], "partyMixId": "0"}]}
sailings = set()
session = requests.Session()
page = session.post(
    "https://disneycruise.disney.go.com/wam/cruise-sales-service/cruise-listing/?region=INTL&storeId=DCL&view=cruise-listing",
    json=json).json()
print(page)
input("Press any key to continue...")
