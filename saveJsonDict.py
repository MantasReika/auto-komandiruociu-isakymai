import json
from pprint import pprint
# comments for no purpose
# blah
# blah blah
# ccoomments for no purpose
# blagh
# blahh blah
GALININKO_LINKSNIAI = {    
            'a' : 'ą',
            'as' : 'ą',
            'ė' : 'ę',
            'is' : 'į',
            'us' : 'ų',
            'ys' : 'į'}

with open('GALININKO_LINKSNIAI.json', 'w', encoding='utf8') as fp:
    json.dump(GALININKO_LINKSNIAI, fp, ensure_ascii=False, indent=4)
    
MONTHS = {
    "01":"sausio",
    "02":"vasario",
    "03":"kovo",
    "04":"balandžio",
    "05":"gegužės",
    "06":"birželio",
    "07":"liepos",
    "08":"rugpjūčio",
    "09":"rugsėjo",
    "10":"spalio",
    "11":"lapkričio",
    "12":"gruodžio"}

with open('MONTHS.json', 'w', encoding='utf8') as fp:
    json.dump(MONTHS, fp, ensure_ascii=False, indent=4)

GA2 = {
       "PGA":"Centro regiono Panevežio GA",
       "MGA":"Vakarų regiono Mažeikių GA",
       "EGA":"Rytų regiono Elektrėnų GA",
       "VGA":"Rytų regiono Vilniaus GA",
       "GTS":"Gamybos technikos skyriaus",
       "EXP":"Eksporto skyriaus ",
       "ADM":"administracijos darbuotoją"}

with open('GA2.json', 'w', encoding='utf8') as fp:
    json.dump(GA2, fp, ensure_ascii=False, indent=4)

GA1 = {
       "PGA":"Centro regiono Panevežio GA darbuotoją",
       "MGA":"Vakarų regiono Mažeikių GA darbuotoją",
       "EGA":"Rytų regiono Elektrėnų GA darbuotoją",
       "VGA":"Rytų regiono Vilniaus GA darbuotoją",
       "GTS":"Gamybos technikos skyriaus",
       "EXP":"Eksporto skyriaus ",
       "ADM":"administracijos darbuotoją"}

with open('GA1.json', 'w', encoding='utf8') as fp:
    json.dump(GA1, fp, ensure_ascii=False, indent=4)

PROFESIJA = {        
        "izol.":"termoizoliuotoją",
        "past.":"pastolininką",
        "d.v.":"darbų vadovą",
        "p.v.":"projektų vadovą",
        "direkt.":"direktorių",
        "vairuot.":"vairuotoją",
        "vert.":"vertėją",
        "suvir.":"suvirintoją",
        "vadyb.":"vadybininkę",
        "buh.":"buhalterę",
        "tech.":"technikę",
        "vadov.":"vadovą"}

with open('PROFESIJA.json', 'w', encoding='utf8') as fp:
    json.dump(PROFESIJA, fp, ensure_ascii=False, indent=4)

DIENPINIGIAI = {
            "Švedija":["32,5", "65"],
            "Prancūzija":["31,5", "63"],
            "Belgija":["30,5", "61"],
            "Šveicarija":["26,5", "53"],
            "Vokietija":["31", "62"],            
            "Latvija":["22", "44"],
            "Lenkija":["24", "48"],
            "Estija":["23,5", "47"],
            "Olandija":["32", "64"],
            "Danija":["40", "80"],
            "Slovakija":["40", "80"]
            }

with open('DIENPINIGIAI.json', 'w', encoding='utf8') as fp:
    json.dump(DIENPINIGIAI, fp, ensure_ascii=False, indent=4)
    
KEYS = ['#DOKYEAR#',
        '#DOKMONTH#',
        '#DOKDAY#',
        '#DOKNUM#',
        '#GA1#',
        '#GA2#',
        '#CNTRY1#',
        '#CNTRY2#',
        '#KOMANDYEAR#',
        '#KOMANDMONTH#',
        '#KOMANDDAY1#',
        '#KOMANDDAY2#',
        '#PERSONNAMECASED#',
        '#CNTRY3#',
        '#PROFESSION#',
        '#MONEY1#',
        '#MONEY2#']

with open('DocumentPlaceHolders.json', 'w', encoding='utf8') as fp:
    json.dump(KEYS, fp, ensure_ascii=False, indent=4)

INVALID_GA = ["ADM","GTS"]

INVALID_SPEC = ["pavad.","sauga","techn.", "vadov.", "pat"]

COUNTRIES = ["Švedija", "Prancūzija", "Belgija", "Šveicarija", "Vokietija", "Latvija", "Ryga", "Lenkija", "Estija", "Nyderlandai", "Olandija", "Suomija", "Danija", "Slovakija"]

TIKSLAS = "Atlikti paskirtus darbus objekte %s pagal sutartį."
#minDIENPINIGIAI = "28,5 Eurų"
#maxDIENPINIGIAI = "57 Eurų"

##FILE_DIR = os.getcwd()
DOCUMENTS_DIR = "Komandiruočių Įsakymai"
FAILED_DOCUMENTS_DIR = "Negarantuoti Komandiruočių Įsakymai"

##with open('GALININKO_LINKSNIAI.json', 'w', encoding='utf8') as fp:
##    json.dump(GALININKO_LINKSNIAI, fp, ensure_ascii=False, indent=4)
##    
##with open('GALININKO_LINKSNIAI.json', 'r', encoding='utf8') as fp:
##    data = json.load(fp)
##    pprint(data, width=1)
