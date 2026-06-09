from flask import Flask, request, jsonify, send_file, render_template_string
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import os
import io
import re
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB limit

# ─────────────────────────────────────────────────────────────────────
# PORT CODE MAP
# ─────────────────────────────────────────────────────────────────────
PORT_CODE_MAP = {'AJMAN': 'AEAJM', 'ABU DHABI': 'AEAUH', 'DUBAI': 'AEDXB', 'AL - FUJAYRAH': 'AEFJR', 'JEBEL ALI': 'AEJEA', 'KHOR AL FAKKAN': 'AEKLF', 'PORT RASHID': 'AEPRA', 'UMM AL QAIWAIN': 'AEQIW', 'RAS AL KHAIMAH': 'AERKT', 'SHARJAH': 'AESHJ', 'YAS ISLAND': 'AEYAS', 'ANTIGUA': 'AGANU', 'ROAD BAY': 'AIRBY', 'DURRES': 'ALDRZ', 'SARANDE': 'ALSAR', 'LUANDA': 'AOLAD', 'LOBITO': 'AOLOB', 'NAMIBE': 'AOMSZ', 'BUENOS AIRES': 'ARBUE', 'CAMPANA': 'ARCMP', 'LA PLATA': 'ARLPG', 'PUERTO MADRYN': 'ARPMY', 'ROSARIO': 'ARROS', 'USHUAIA': 'ARUSH', 'ZARATE': 'ARZAE', 'PAGO PAGO': 'ASPPG', 'ABBOT POINT': 'AUABP', 'ADELAIDE': 'AUADL', 'ALBANY': 'AUALH', 'BELL BAY': 'AUBEL', 'BRISBANE': 'AUBNE', 'BOOBY ISLAND': 'AUBOO', 'CAIRNS': 'AUCNS', 'DARWIN': 'AUDRW', 'ESPERANCE': 'AUEPR', 'FREMANTLE': 'AUFRE', 'GLADSTONE': 'AUGLT', 'HOBART': 'AUHBA', 'HAY POINT': 'AUHPT', 'MELBOURNE': 'AUMEL', 'NEWCASTLE': 'GBNCS', 'PORT HEDLAND': 'AUPHE', 'PORT KEMBLA': 'AUPKL', 'SYDNEY': 'CASYD', 'ORANJESTAD': 'AWORJ', 'BRIDGETOWN': 'BBBGI', 'CHATTOGRAM': 'BDCGP', 'MONGLA': 'BDMGL', 'ANTWERP': 'BEANR', 'GENT (GHENT)': 'BEGNE', 'ZEEBRUGGE': 'BEZEE', 'BURGAS': 'BGBOJ', 'VARNA': 'BGVAR', 'BAHRAIN': 'BHKBS', 'COTONOU': 'BJCOO', 'HAMILTON': 'BMBDA', 'KINGS WHARF': 'BMKWF', 'SINT EUSTATIUS': 'BQEUX', 'ARRAIAL DO CABO': 'BRACB', 'ANGRA DOS REIS': 'BRADR', 'ARTUR NOGUEIRA': 'BRANG', 'BELEM': 'BRBEL', 'BUZIOS': 'BRBZC', 'CAMBORIU': 'BRCBU', 'CABEDELO': 'BRCDO', 'CARMO DO PARANAIBA': 'BRCOP', 'FORTALEZA': 'BRFOR', 'IMBITUBA': 'BRIBB', 'ILHABELA': 'BRIBE', 'ILHA GRANDE': 'BRIGE', 'ITAGUAI': 'BRIGI', 'ITAPOA': 'BRIOA', 'ILHEUS': 'BRIOS', 'ITAJAI': 'BRITJ', 'MANAUS': 'BRMAO', 'MACEIO': 'BRMCZ', 'NAVEGANTES': 'BRNVT', 'PORTO BELO': 'BRPBO', 'PECEM': 'BRPEC', 'PARANAGUA': 'BRPNG', 'PORTO VELHO': 'BRPVH', 'CABO FRIO': 'BRQCK', 'RECIFE': 'BRREC', 'RIO GRANDE': 'BRRIG', 'RIO DE JANEIRO': 'BRRIO', 'SAO FRANCISCO DO SUL': 'BRSFS', 'SALVADOR': 'BRSSA', 'SANTOS': 'BRSSZ', 'SANTAREM': 'BRSTM', 'SUAPE': 'BRSUA', 'UBATUBA': 'BRUBT', 'VITORIA': 'BRVIX', 'VILA DO CONDE': 'BRVLC', 'COCO CAY': 'BSCOC', 'FREEPORT, GRAND BAHAMA': 'BSFPO', 'GREAT STIRRUP CAY': 'BSGSC', 'LITTLE SAN SALVADOR': 'BSHMC', 'LUCAYA': 'BSLUC', 'NASSAU': 'BSNAS', 'BELIZE CITY': 'BZBZE', 'BAIE COMEAU': 'CABCO', 'BECANCOUR': 'CABEC', 'CORNER BROOK': 'CACBK', 'CHARLOTTETOWN': 'CACHA', 'GASPE': 'CAGPE', 'HALIFAX': 'CAHAL', 'HAVRE-SAINT-PIERRE': 'CAHSP', 'LA BAIE': 'CALBA', 'MONTREAL': 'CAMTR', 'PRINCE RUPERT': 'CAPRR', 'QUEBEC': 'CAQUE', 'SAINT JOHN': 'CASJB', 'TROIS-RIVIERES (THREE RIVERS)': 'CATRR', 'VANCOUVER': 'CAVAN', 'COCOS ISLANDS': 'CCCCK', 'MATADI': 'CDMAT', 'POINTE NOIRE': 'CGPNR', 'ABIDJAN': 'CIABJ', 'SAN-PEDRO': 'CISPY', 'AITUTAKI': 'CKAIT', 'RAROTONGA': 'CKRAR', 'ANTOFAGASTA': 'CLANF', 'ARICA': 'CLARI', 'CORONEL': 'CLCNL', 'COQUIMBO': 'CLCQQ', 'ISLA DE PASCUA': 'CLIPC', 'IQUIQUE': 'CLIQQ', 'LIRQUEN': 'CLLQN', 'PUERTO ANGAMOS': 'CLPAG', 'MAGELLAN STRAIT': 'CLPPY', 'SAN ANTONIO': 'CLSAI', 'SAN VICENTE': 'CLSVE', 'VALPARAISO': 'CLVAP', 'DOUALA': 'CMDLA', 'TIESHAN': 'CNBHT', 'GUANGZHOU': 'CNCAN', 'CAOFEIDIAN': 'CNCFD', 'DA CHAN BAY': 'CNDCB', 'DALIAN': 'CNDLC', 'FANGCHENG': 'CNFAN', 'FUZHOU': 'CNFOC', 'HAIKOU': 'CNHAK', 'HUANGHUA': 'CNHNH', 'HUANGPU': 'CNHUA', 'JINGTANG (TANGSHAN)': 'CNJIN', 'LU-HUA SHAN': 'CNLUH', 'LIANYUNGANG': 'CNLYG', 'MAWEI': 'CNMAW', 'MAWAN': 'CNMWN', 'NINGBO': 'CNNGB', 'NANJING': 'CNNKG', 'NANSHA': 'CNNSA', 'NANTONG': 'CNNTG', 'QINZHOU': 'CNQZH', 'SHANGHAI': 'CNSHA', 'SANSHAN': 'CNSHG', 'SHEKOU': 'CNSHK', 'QINHUANGDAO': 'CNSHP', 'SHANTOU': 'CNSWA', 'TAICANG': 'CNTAG', 'QINGDAO': 'CNTAO', 'TIANJINXINGANG': 'CNTXG', 'WEIHAI': 'CNWEI', 'XIAMEN': 'CNXMN', 'XIANGSHUI': 'CNXSI', 'YANTAI': 'CNYNT', 'YANTIAN': 'CNYTN', 'ZHANGJIAGANG': 'CNZJG', 'ZHOUSHAN': 'CNZOS', 'BARRANQUILLA': 'COBAQ', 'BUENAVENTURA': 'COBUN', 'CARTAGENA': 'ESCAR', 'TURBO': 'COTRB', 'CALDERA': 'CRCAL', 'PUERTO LIMON': 'CRLIO', 'MOIN': 'CRMOB', 'NUEVA GERONA': 'CUGER', 'LA HABANA': 'CUHAV', 'MARIEL': 'CUMAR', 'MINDELO': 'CVMIN', 'PRAIA': 'CVRAI', 'SAL REI': 'CVSAR', 'CURACAO': 'CWCUR', 'FAMAGUSTA': 'CYFMG', 'LIMASSOL': 'CYLMS', 'BRUNSBUTTEL': 'DEBRB', 'BREMEN': 'DEBRE', 'BREMERHAVEN': 'DEBRV', 'ELSFLETH': 'DEELS', 'EMDEN': 'DEEME', 'HAMBURG': 'DEHAM', 'HELGOLAND': 'DEHGL', 'KIEL': 'DEKEL', 'LUBECK': 'DELBC', 'WARNEMUNDE': 'DEWAR', 'WILHELMSHAVEN': 'DEWVN', 'DJIBOUTI': 'DJJIB', 'AALBORG': 'DKAAL', 'AARHUS': 'DKAAR', 'COPENHAGEN': 'DKCPH', 'FREDERICIA': 'DKFRC', 'GREAT BELT': 'DKGBT', 'HVIDE SANDE': 'DKHVS', 'KALUNDBORG': 'DKKAL', 'KOLIND': 'DKKLD', 'KERTEMINDE': 'DKKTD', 'ODENSE': 'DKODE', 'RONNE': 'DKRNN', 'SKAGEN': 'DKSKA', 'PORTSMOUTH': 'USPTM', 'ROSEAU': 'DMRSU', 'CAUCEDO': 'DOCAU', 'RIO HAINA': 'DOHAI', 'PUERTO PLATA': 'DOPOP', 'ANNABA': 'DZAAE', 'ALGER': 'DZALG', 'ARZEW': 'DZAZW', 'BEJAIA': 'DZBJA', 'DJEN-DJEN': 'DZDJE', 'GHAZAOUET': 'DZGHZ', 'ORAN': 'DZORN', 'SKIKDA': 'DZSKI', 'ESMERALDAS': 'ECESM', 'GUAYAQUIL': 'ECGYE', 'LA LIBERTAD': 'ECLLD', 'PUERTO BOLIVAR': 'ECPBO', 'POSORJA': 'ECPSJ', 'TALLINN': 'EETLL', 'ALEXANDRIA OLD PORT': 'EGALY', 'DAMIETTA': 'EGDAM', 'ALEXANDRIA EL DEKHEILA': 'EGEDK', 'PORT SAID EAST': 'EGPSE', 'PORT SAID WEST': 'EGPSW', 'SAFAGA': 'EGSGA', 'SOKHNA PORT': 'EGSOK', 'SHARM ASH SHAYKH': 'EGSSH', 'SUEZ': 'EGSUZ', 'ASSAB': 'ERASA', 'ARRECIFE DE LANZAROTE': 'ESACE', 'MALAGA': 'ESAGP', 'ALICANTE': 'ESALC', 'ALGECIRAS': 'ESALG', 'ALMAGRO': 'ESALM', 'BARCELONA': 'ESBCN', 'BILBAO': 'ESBIO', 'CADIZ': 'ESCAD', 'CASTELLON DE LA PLANA': 'ESCAS', 'CEUTA': 'ESCEU', 'FERROL': 'ESFRO', 'PUERTO DEL ROSARIO-FUERTEVENTURA': 'ESFUE', 'GIJON': 'ESGIJ', 'GRANJA DE SAN IDELFONSO': 'ESGJI', 'HUELVA': 'ESHUV', 'IBIZA': 'ESIBZ', 'LA CORUNA': 'ESLCG', 'ALMERIA': 'ESLEI', 'LAS PALMAS': 'ESLPA', 'MAHON, MENORCA': 'ESMAH', 'MARIN, PONTEVEDRA': 'ESMPG', 'PALAMOS': 'ESPAL', 'PALMA DE MALLORCA': 'ESPMI', 'ROSAS': 'ESROS', 'SAGUNTO': 'ESSAG', 'SANTA CRUZ DE TENERIFE': 'ESSCT', 'SANTA CRUZ DE LA PALMA': 'ESSPC', 'SAN SEBASTIAN DE LA GOMERA': 'ESSSG', 'SEVILLA': 'ESSVQ', 'TARRAGONA': 'ESTAR', 'VIGO': 'ESVGO', 'VALENCIA': 'ESVLC', 'HELSINKI': 'FIHEL', 'KEMI': 'FIKEM', 'KOKKOLA (KARLEBY)': 'FIKOK', 'KOTKA': 'FIKTK', 'OULU (ULEABORG)': 'FIOUL', 'RAUMA': 'FIRAU', 'TORNIO (TORNEA)': 'FITOR', 'LAUTOKA': 'FJLTK', 'SUVA': 'FJSUV', 'THORSHAVN': 'FOTHO', 'AJACCIO': 'FRAJA', 'BREST': 'FRBES', 'BORDEAUX': 'FRBOD', 'CANNES': 'FRCEQ', 'CHERBOURG': 'FRCER', 'DUNKERQUE': 'FRDKK', 'FOS-SUR-MER': 'FRFOS', 'GENNEVILLIERS': 'FRGVL', 'HONFLEUR': 'FRHON', 'LE HAVRE': 'FRLEH', 'LA ROCHELLE': 'FRLRH', 'LE VERDON': 'FRLVE', 'MARSEILLE': 'FRMRS', 'MONTOIR-DE-BRETAGNE': 'FRMTX', 'NICE': 'FRNCE', 'SETE': 'FRSET', 'SIX-FOURS-LES-PLAGES': 'FRSFP', 'ST NAZAIRE': 'FRSNR', 'ST MARCEL': 'FRSTM', 'ST TROPEZ': 'FRSTP', 'TOULON': 'FRTLN', 'ROUEN': 'FRURO', 'VILLEFRANCHE-SUR-MER': 'FRVFM', 'LIBREVILLE': 'GALBV', 'PORT GENTIL': 'GAPOG', 'BELFAST': 'GBBEL', 'BRISTOL': 'GBBRS', 'DOVER': 'GBDVR', 'FALMOUTH': 'JMFMH', 'FELIXSTOWE': 'GBFXT', 'GRANGEMOUTH': 'GBGRG', 'GREENOCK': 'GBGRK', 'HARWICH': 'GBHRW', 'HULL': 'GBHUL', 'IMMINGHAM': 'GBIMM', 'LERWICK': 'GBLER', 'LONDON GATEWAY PORT': 'GBLGP', 'LIVERPOOL': 'GBLIV', 'PORTREE': 'GBPRT', 'PORTBURY': 'GBPRU', 'PORTLAND': 'GBPTL', 'SOUTH QUEENSFERRY': 'GBSOQ', 'SOUTHAMPTON': 'GBSOU', 'SOUTH SHIELDS': 'GBSSH', 'TEESPORT': 'GBTEE', 'THAMESPORT': 'GBTHP', 'TILBURY': 'GBTIL', 'GRENADA': 'GDGND', "SAINT GEORGE'S": 'GDSTG', 'BATUMI': 'GEBUS', 'POTI': 'GEPTI', 'ST PETER PORT': 'GGSPT', 'TEMA': 'GHTEM', 'TAKORADI': 'GHTKD', 'GIBRALTAR': 'GIGIB', 'NUUK (GODTHAAB)': 'GLGOH', 'PAAMIUT (FREDRIKSHAAB)': 'GLJFR', 'BANJUL': 'GMBJL', 'CONAKRY': 'GNCKY', 'POINTE-A-PITRE': 'GPPTP', 'ARGOSTOLION': 'GRARM', 'KERKIRA (CORFU)': 'GRCFU', 'CANEA (CHANIA)': 'GRCHQ', 'ELEFSIS (ELEVSIS)': 'GRELE', 'GYTHION': 'GRGYT', 'HERAKLION': 'GRHER', 'MYKONOS': 'GRJMK', 'SYROS (SYRA)': 'GRJSY', 'THIRA': 'GRJTR', 'KATAKOLON': 'GRKAK', 'KOS': 'GRKGS', 'KALILIMENES': 'GRKLL', 'KALAMATA': 'GRKLX', 'LAURIUM (LAVRION)': 'GRLAV', 'MOUDHROS': 'GRMDR', 'MONEMVASIA': 'GRMON', 'NAFPLION': 'GRNAF', 'PIRAEUS': 'GRPIR', 'PATMOS': 'GRPMS', 'RHODES': 'GRRHO', 'SOUDA': 'GRSDH', 'SKARAMANGAS': 'GRSKA', 'THESSALONIKI': 'GRSKG', 'TILOS': 'GRTIL', 'VOLOS': 'GRVOL', 'ZAKYNTHOS': 'GRZTH', 'PUERTO BARRIOS': 'GTPBR', 'PUERTO QUETZAL': 'GTPRQ', 'PUERTO SANTO TOMAS DE CASTILLA': 'GTSTC', 'BISSAU': 'GWOXB', 'GEORGETOWN': 'GYGEO', 'HONG KONG': 'HKHKG', 'PUERTO CASTILLA': 'HNPCA', 'PUERTO CORTES': 'HNPCR', 'ROATAN': 'HNRTB', 'SAN LORENZO': 'HNSLO', 'DUBROVNIK': 'HRDBV', 'PLOCE': 'HRPLE', 'RIJEKA': 'HRRJK', 'SPLIT': 'HRSPU', 'ZADAR': 'HRZAD', 'GONAIVES': 'HTGVS', 'LABADIE': 'HTLAB', 'PORT AU PRINCE': 'HTPAP', 'BELAWAN, SUMATRA': 'IDBLW', 'BENOA, BALI': 'IDBOA', 'BALIKPAPAN, KALIMANTAN': 'IDBPN', 'BATAM ISLAND': 'IDBTM', 'JAMBI, SUMATRA': 'IDDJB', 'JAKARTA, JAVA': 'IDJKT', 'MAKASSAR': 'IDMAK', 'MANGOLE': 'IDMAL', 'PADANG': 'IDPDG', 'PERAWANG': 'IDPER', 'PALEMBANG, SUMATRA': 'IDPLM', 'PANJANG': 'IDPNJ', 'PONTIANAK, KALIMANTAN': 'IDPNK', 'PERAWANG, SUMATRA': 'IDPWG', 'SEMARANG': 'IDSRG', 'SURABAYA': 'IDSUB', 'TABONEO': 'IDTAB', 'TANJUNG BARA, KL': 'IDTBA', 'UJUNG PANDANG, SULAWESI': 'IDUPG', 'DUN LAOGHAIRE': 'IEDLG', 'DUBLIN': 'IEDUB', 'GREENCASTLE': 'IEGRE', 'CORK': 'IEORK', 'WATERFORD': 'IEWAT', 'ASHDOD': 'ILASH', 'ELAT (EILATH)': 'ILETH', 'HAIFA': 'ILHFA', 'ALANG SBY': 'INALA', 'BHAVNAGAR': 'INBHU', 'MUMBAI': 'INBOM', 'KOLKATA': 'INCCU', 'COCHIN': 'INCOK', 'ENNORE': 'INENR', 'HALDIA': 'INHAL', 'HAZIRA PORT/SURAT': 'INHZA', 'KANDLA': 'INIXY', 'KAKINADA': 'INKAK', 'KATTUPALLI': 'INKAT', 'KRISHNAPATNAM': 'INKRI', 'CHENNAI': 'INMAA', 'MARMUGAO (MARMAGAO)': 'INMRM', 'MUNDRA': 'INMUN', 'NEW MANGALORE': 'INNML', 'NHAVA SHEVA': 'INNSA', 'PIPAVAV (VICTOR) PORT': 'INPAV', 'PARADIP GARH': 'INPRT', 'TUTICORIN': 'INTUT', 'VISAKHAPATNAM': 'INVTZ', 'UMM QASR PT': 'IQUQR', 'AKUREYRI': 'ISAKU', 'ISAFJORDUR - HOFN': 'ISISA', 'REYKJAVIK': 'ISREY', 'ALGHERO': 'ITAHO', 'ANCONA': 'ITAOI', 'AUGUSTA PORT OF CATANIA': 'ITAUG', 'BRINDISI': 'ITBDS', 'BARI': 'ITBRI', 'CAGLIARI': 'ITCAG', 'CATANIA': 'ITCTA', 'CIVITAVECCHIA': 'ITCVV', 'GIOIA TAURO': 'ITGIT', 'GENOA': 'ITGOA', 'ISOLA SANTO STEFANO': 'ITISS', 'LEGHORN': 'ITLIV', 'MONFALCONE': 'ITMNF', 'MESSINA': 'ITMSN', 'NAPLES': 'ITNAP', 'OLBIA': 'ITOLB', 'PALAZZOLO DELLO STELLA': 'ITPAL', 'MARGHERA': 'ITPMA', 'PALERMO': 'ITPMO', 'PORTOPALO': 'ITPPL', 'PORTOFINO': 'ITPTF', 'POZZALLO': 'ITPZL', 'SASSARI': 'ITQSS', 'RAVENNA': 'ITRAN', 'SORRENTO': 'ITRRO', 'SALERNO': 'ITSAL', 'SIRACUSA': 'ITSIR', 'LA SPEZIA': 'ITSPE', 'SAVONA': 'ITSVN', 'TARANTO': 'ITTAR', 'TRAPANI': 'ITTPS', 'TRIESTE': 'ITTRS', 'VENICE': 'ITVCE', 'VADO LIGURE': 'ITVDL', 'KINGSTON': 'JMKIN', 'MONTEGO BAY': 'JMMBJ', 'OCHO RIOS': 'JMOCJ', "AL 'AQABAH": 'JOAQJ', 'ABURATSU': 'JPABU', 'AOMORI': 'JPAOJ', 'AKITA': 'JPAXT', 'BEPPU, SHIMANE': 'JPBEP', 'CHIBA': 'JPCHB', 'FUKUYAMA, HIROSHIMA': 'JPFKY', 'FUKUOKA': 'JPFUK', 'HIBIKISHINKO': 'JPHBK', 'HACHINOHE, AOMORI': 'JPHHE', 'HITACHINAKA': 'JPHIC', 'HIROSHIMA': 'JPHIJ', 'HIMEJI': 'JPHIM', 'HAKODATE': 'JPHKD', 'HAKATA, FUKUOKA': 'JPHKT', 'HAMADA': 'JPHMD', 'HOSOSHIMA': 'JPHHSM', 'IMABARI': 'JPIMB', 'IMARI': 'JPIMI', 'ISHIKARI': 'JPISI', 'IWAKUNI': 'JPIWK', 'IYOMISHIMA': 'JPIYM', 'KOCHI': 'JPKCZ', 'NIIGATA': 'JPKIJ', 'KAMAISHI': 'JPKIS', 'KUMAMOTO': 'JPKMJ', 'KANAZAWA': 'JPKNZ', 'KAGOSHIMA': 'JPKOJ', 'KURE, HIROSHIMA': 'JPKRE', 'KASHIMA, IBARAKI': 'JPKSM', 'KUSHIRO': 'JPKUH', 'KAWASAKI': 'JPKWS', 'MAIZURU': 'JPMAI', 'MIIKE, FUKUOKA': 'JPMII', 'MIZUSHIMA': 'JPMIZ', 'MOJI/KITAKYUSHU': 'JPMOJ', 'MURORAN': 'JPMUR', 'MATSUYAMA': 'JPMYJ', 'MIYAKO, IWATE': 'JPMYK', 'NAHA, OKINAWA': 'JPNAH', 'NAKANOSEKI': 'JPNAN', 'NAOETSU': 'JPNAO', 'NAGOYA': 'JPNGO', 'NAGASAKI': 'JPNGS', 'OHFUNATO': 'JPOFT', 'OITA': 'JPOIT', 'OMAEZAKI': 'JPOMZ', 'ONAHAMA': 'JPONA', 'OSAKA': 'JPOSA', 'OTAKE': 'JPOTK', 'SHIBUSHI': 'JPSBS', 'SENDAI, MIYAGI': 'JPSDJ', 'SHIMONOSEKI': 'JPSHS', 'SAKATA': 'JPSKT', 'SAKAIMINATO': 'JPSMN', 'SHIMIZU': 'JPSMZ', 'SATSUMASENDAI': 'JPSTS', 'TAKAMATSU': 'JPTAK', 'TOYOHASHI': 'JPTHS', 'TOKUSHIMA': 'JPTKS', 'TOKUYAMA': 'JPTKY', 'TOMAKOMAI': 'JPTMK', 'TOYAMASHINKO': 'JPTOS', 'TSURUGA': 'JPTRG', 'TOKYO': 'JPTYO', 'UBE': 'JPUBJ', 'KOBE': 'JPUKB', 'WAKAYAMA': 'JPWAK', 'YATSUSHIRO': 'JPYAT', 'YOKKAICHI': 'JPYKK', 'YOKOHAMA': 'JPYOK', 'EMBAKASI': 'KEEMB', 'MOMBASA': 'KEMBA', 'KAMPONG SAOM (SIHANOUKVILLE)': 'KHKOS', 'PHNOM PENH': 'KHPNH', 'MUTSAMUDU': 'KMMUT', 'MORONI': 'KMYVA', 'BASSETERRE, ST KITTS': 'KNBAS', 'NEVIS': 'KNNEV', 'GOSEONG-GUN': 'KRGSO', 'INCHEON': 'KRINC', 'GANGNEUNG': 'KRKAG', 'GWANGYANG': 'KRKAN', 'POHANG': 'KRKPO', 'MOKPO': 'KRMOK', 'OKPO/GEOJE': 'KROKP', 'BUSAN': 'KRPUS', 'SAMCHEONPO/SACHEON': 'KRSCP', 'SEOGWIPO': 'KRSPO', 'DANGJIN': 'KRTJI', 'TONGYEONG': 'KRTYG', 'ULSAN': 'KRUSN', 'YEOSU': 'KRYOS', 'SHUAIBA': 'KWSAA', 'SHUWAIKH': 'KWSWK', 'GRAND CAYMAN': 'KYGCM', 'ALMATY': 'KZALA', 'BEIRUT': 'LBBEY', 'CASTRIES': 'LCCAS', 'ST LUCIA APT': 'LCSLU', 'COLOMBO': 'LKCMB', 'GALLE': 'LKGAL', 'HAMBANTOTA': 'LKHBA', 'TRINCOMALEE': 'LKTRR', 'MONROVIA': 'LRMLW', 'KLAIPEDA': 'LTKLJ', 'RIGA': 'LVRIX', 'BINGAZI': 'LYBEN', 'KHOMS': 'LYKHO', 'MISURATA': 'LYMRA', 'TRIPOLI': 'LYTIP', 'AGADIR': 'MAAGA', 'CASABLANCA': 'MACAS', 'NADOR': 'MANDR', 'TANGER MED': 'MAPTM', 'MONTE-CARLO': 'MCMCM', 'GIURGIULESTI': 'MDGIU', 'BAR': 'MEBAR', 'BIJELA': 'MEBIJ', 'KOTOR': 'MEKOT', 'DIEGO SUAREZ': 'MGDIE', 'EHOALA (TOLAGNARO)': 'MGEHL', 'FORT DAUPHIN (TOALAGNARO)': 'MGFTU', 'MAJUNGA': 'MGMJN', 'NOSY-BE': 'MGNOS', 'SAINTE MARIE': 'MGSMS', 'TULEAR': 'MGTLE', 'TAMATAVE': 'MGTMM', 'VOHEMAR': 'MGVOH', 'MANAKARA': 'MGWVK', 'YANGON': 'MMRGN', 'MACAU': 'MOMFM', 'FORT-DE-FRANCE': 'MQFDF', 'NOUADHIBOU': 'MRNDB', 'NOUAKCHOTT': 'MRNKC', 'LITTLE BAY': 'MSLTB', 'PLYMOUTH': 'MSPLY', 'MARSAXLOKK': 'MTMAR', 'VALLETTA': 'MTMLA', 'PORT LOUIS': 'MUPLU', 'MALE': 'MVMLE', 'ALTAMIRA': 'MXATM', 'COSTA MAYA': 'MXCOM', 'COZUMEL': 'MXCZM', 'ENSENADA': 'MXESE', 'GUAYMAS': 'MXGYM', 'LAZARO CARDENAS': 'MXLZC', 'MAZATLAN': 'MXMZT', 'PROGRESO': 'MXPGO', 'PUERTO MORELOS': 'MXPMS', 'TAMPICO': 'MXTAM', 'TULUM': 'MXTUY', 'VERACRUZ': 'MXVER', 'MANZANILLO': 'PAMIT', 'KOTA KINABALU, SABAH': 'MYBKI', 'BINTULU, SARAWAK': 'MYBTU', 'KUCHING, SARAWAK': 'MYKCH', 'KUANTAN': 'MYKUA', 'LABUAN, SABAH': 'MYLBU', 'LANGKAWI': 'MYLGK', 'MALACCA': 'MYMKZ', 'MIRI, SARAWAK': 'MYMYY', 'PENANG': 'MYPEN', 'PASIR GUDANG, JOHOR': 'MYPGU', 'PORT KLANG (PELABUHAN KLANG)': 'MYPKG', 'SIBU, SARAWAK': 'MYSBW', 'SANDAKAN, SABAH': 'MYSDK', 'TANJUNG PELEPAS': 'MYTPP', 'TAWAU, SABAH': 'MYTWU', 'BEIRA': 'MZBEW', 'BAZARUTO ISLAND': 'MZBZB', 'NACALA': 'MZMNC', 'MAPUTO': 'MZMPM', 'MASSINGA': 'MZMSG', 'PEMBA': 'MZPOL', 'QUELIMANE': 'MZUEL', 'LUDERITZ': 'NALUD', 'WALVIS BAY': 'NAWVB', 'NOUMEA': 'NCNOU', 'VAVOUTO': 'NCVAV', 'APAPA': 'NGAPP', 'LEKKI': 'NGLKK', 'ONNE': 'NGONN', 'PORT HARCOURT': 'NGPHC', 'TINCAN/LAGOS': 'NGTIN', 'CORINTO': 'NICIO', 'MANAGUA': 'NIMGA', 'AMSTERDAM': 'NLAMS', 'FLUSHING': 'NLFLU', 'HEERENVEEN': 'NLHRV', 'IJMUIDEN': 'NLIJM', 'MOERDIJK': 'NLMOE', 'ROTTERDAM': 'NLRTM', 'VLISSINGEN': 'NLVLI', 'ALESUND': 'NOAES', 'ALTA': 'NOALF', 'ALSTAHAUG': 'NOALS', 'AUSTEVOLL': 'NOASV', 'BERGEN': 'NOBGO', 'BREVIK': 'NOBVK', 'EIDFJORD': 'NOEDF', 'EGERSUND': 'NOEGE', 'FLAM': 'NOFLA', 'FREDRIKSTAD': 'NOFRK', 'FLORO': 'NOFRO', 'GJEMNES': 'NOGJM', 'GEIRANGER': 'NOGNR', 'HALDEN': 'NOHAL', 'HAUGESUND': 'NOHAU', 'HOGSET': 'NOHOG', 'HONNINGSVAG': 'NOHVG', 'IKORNNES': 'NOIKR', 'KARMOY': 'NOKMY', 'KRISTIANSAND': 'NOKRS', 'KRISTIANSUND': 'NOKSU', 'KVINESDAL': 'NOKVD', 'LARVIK': 'NOLAR', 'LEIKANGER': 'NOLEK', 'LEKNES': 'NOLKN', 'LONGYEARBYEN': 'NOLYR', 'MALOY': 'NOMAY', 'MOLDE': 'NOMOL', 'MOSS': 'NOMSS', 'NARVIK': 'NONVK', 'OLDEN': 'NOOLD', 'ORKANGER': 'NOORK', 'OSLO': 'NOOSL', 'SALTEN': 'NOSAT', 'SAUDA': 'NOSAU', 'SUNNDALSORA': 'NOSUN', 'SVELGEN': 'NOSVE', 'STAVANGER': 'NOSVG', 'TANANGER': 'NOTAE', 'TROMSO': 'NOTOS', 'TRONDHEIM': 'NOTRD', 'AUCKLAND': 'NZAKL', 'BLUFF': 'NZBLU', 'CHRISTCHURCH': 'NZCHC', 'DUNEDIN': 'NZDUD', 'LYTTELTON': 'NZLYT', 'MARSDEN POINT': 'NZMAP', 'NAPIER': 'NZNPE', 'NELSON': 'NZNSN', 'PORT CHALMERS': 'NZPOE', 'TIMARU': 'NZTIU', 'TAURANGA': 'NZTRG', 'WELLINGTON': 'NZWLG', 'DUQM': 'OMDQM', 'KHASAB': 'OMKHS', 'MUSCAT': 'OMMCT', 'SALALAH': 'OMSLL', 'SOHAR': 'OMSOH', 'BALBOA': 'PABLB', 'CRISTOBAL': 'PACTB', 'COLON': 'PAONX', 'ALMIRANTE': 'PAPAM', 'PANAMA': 'PAPTY', 'RODMAN': 'PAROD', 'CALLAO': 'PECLL', 'PAITA': 'PEPAI', 'PISCO': 'PEPIO', 'SALAVERRY': 'PESVY', 'BORA-BORA': 'PFBOB', 'MOOREA': 'PFMOZ', 'PAPEETE': 'PFPPT', 'ALOTAU': 'PGGUR', 'LAE': 'PGLAE', 'LOSUIA': 'PGLSA', 'PORT MORESBY': 'PGPOM', 'RABAUL': 'PGRAB', 'BATANGAS, LUZON': 'PHBTG', 'CEBU': 'PHCEB', 'CAGAYAN DE ORO, MINDANAO': 'PHCGY', 'DAVAO, MINDANAO': 'PHDVO', 'GENERAL SANTOS': 'PHGES', 'MANILA NORTH HARBOUR': 'PHMNN', 'MANILA SOUTH HARBOUR': 'PHMNS', 'SUBIC': 'PHSFS', 'KARACHI-MUHAMMAD BIN QASIM': 'PKBQM', 'KARACHI': 'PKKHI', 'GDANSK': 'PLGDN', 'GDYNIA': 'PLGDY', 'SWINOUJSCIE': 'PLSWI', 'SZCZECIN': 'PLSZZ', 'PITCAIRN IS': 'PNPCN', 'SAN JUAN': 'PRSJU', 'AVEIRO': 'PTAVE', 'FIGUEIRA DA FOZ': 'PTFDF', 'FUNCHAL, MADEIRA': 'PTFNC', 'LEIXOES': 'PTLEI', 'LISBOA': 'PTLIS', 'PONTA DELGADA': 'PTPDL', 'PORTIMAO': 'PTPRM', 'SETUBAL': 'PTSET', 'SINES': 'PTSIE', 'SESIMBRA': 'PTSSB', 'TERCEIRA ISLAND': 'PTTER', 'KOROR': 'PWROR', 'CAACUPEMI ASUNCION': 'PYBCM', 'ENCARNACION PUERTO SAN JUAN': 'PYENO', 'CAACUPEMI PILAR': 'PYPIL', 'TERPORT VILLETA': 'PYTVT', 'PUERTO SEGURO FLUVIAL (VILLETA)': 'PYVLL', 'DOHA': 'QADOH', 'HAMAD': 'QAHMD', 'MESAIEED': 'QAMES', 'RAS LAFFAN': 'QARLF', 'POINTE DES GALETS': 'REPDG', 'POSSESSION': 'REPOS', 'AGIGEA': 'ROAGI', 'CONSTANTA': 'ROCND', 'GALATI': 'ROGAL', 'MANGALIA': 'ROMAG', 'ARKHANGELSK': 'RUARH', 'BALTIYSK': 'RUBLT', 'KRONSHTADT': 'RUKDT', 'KAVKAZ': 'RUKZP', 'SAINT PETERSBURG': 'RULED', 'NAKHODKA': 'RUNJK', 'NOVOROSSIYSK': 'RUNVS', 'PETROPAVLOVSK-KAMCHATSKIY': 'RUPKC', 'SLAVYANKA': 'RUSKA', 'SOCHI': 'RUSOC', "UST'-LUGA": 'RUULU', 'VLADIVOSTOK': 'RUVVO', 'ZARUBINO': 'RUZAR', 'AD DAMMAM': 'SADMM', 'JEDDAH': 'SAJED', 'JUBAIL': 'SAJUB', 'KING ABDULLAH PORT': 'SAKAC', 'NEOM': 'SANEO', 'YANBU AL-BAHR': 'SAYNB', 'HONIARA, GUADALCANAL IS': 'SBHIR', 'PORT VICTORIA': 'SCPOV', 'VICTORIA': 'SCVIC', 'PORT SUDAN': 'SDPZU', 'AHUS': 'SEAHU', 'GOTEBORG': 'SEGOT', 'GAVLE': 'SEGVX', 'HALMSTAD': 'SEHAD', 'HELSINGBORG': 'SEHEL', 'KARLSHAMN': 'SEKAN', 'NORRKOPING': 'SENRK', 'NYNASHAMN': 'SENYN', 'PITEA': 'SEPIT', 'SKELLEFTEA': 'SESFT', 'SKARHAMN': 'SESKM', 'SODERTALJE': 'SESOE', 'STOCKHOLM': 'SESTO', 'VISBY': 'SEVBY', 'SINGAPORE': 'SGSIN', 'JAMESTOWN': 'SHSHN', 'KOPER': 'SIKOP', 'FREETOWN': 'SLFNA', 'DAKAR': 'SNDKR', 'ZIGUINCHOR': 'SNZIG', 'BERBERA': 'SOBBO', 'KISMAYU': 'SOKMU', 'MOGADISHU': 'SOMGQ', 'PARAMARIBO': 'SRPBM', 'ACAJUTLA': 'SVAQJ', 'PHILIPSBURG': 'SXPHI', 'LATTAKIA': 'SYLTK', 'TARTUS': 'SYTTS', 'GRAND TURK ISLAND': 'TCGDT', 'PROVIDENCIALES': 'TCPLS', 'LOME': 'TGLFW', 'BANGKOK': 'THBKK', 'BANGKOK MODERN TERMINALS/BANGKOK': 'THBMT', 'PHUKET': 'THHKT', 'LAEM CHABANG': 'THLCH', 'LAT KRABANG': 'THLKR', 'PAT BANGKOK': 'THPAT', 'SIAM BANGKOK PORT': 'THSBP', 'SONGKHLA': 'THSGZ', 'THAI CONNECTIVITY TERMINAL': 'THTPT', 'KOH SAMUI': 'THUSM', 'DILI': 'TLDIL', 'LA GOULETTE NORD (HALQUELOUED)': 'TNLGN', 'RADES/TUNIS': 'TNRDS', 'SFAX': 'TNSFA', 'SOUSSE': 'TNSUS', 'TUNIS': 'TNTUN', "NUKU'ALOFA": 'TOTBU', 'ALANYA': 'TRALA', 'ALIAGA': 'TRALI', 'AVCILAR': 'TRAVC', 'ANTALYA': 'TRAYT', 'BANDIRMA': 'TRBDM', 'BESIKTAS': 'TRBTS', 'BODRUM': 'TRBXN', 'BOZCAADA': 'TRBZC', 'CANAKKALE': 'TRCKZ', 'DERINCE': 'TRDRC', 'EVYAP PORT': 'TREYP', 'GEBZE': 'TRGEB', 'GEMLIK': 'TRGEM', 'GIRESUN': 'TRGIR', 'ISKENDERUN': 'TRISK', 'ISTANBUL': 'TRIST', 'ISTINYE/BOSPHORUS': 'TRITY', 'IZMIR': 'TRIZM', 'LIMAS': 'TRLMA', 'MERSIN': 'TRMER', 'MARMARIS': 'TRMRM', 'SAMSUN': 'TRSSX', 'TEKIRDAG (ASYAPORT)': 'TRTEK', 'TUZLA': 'TRTUZ', 'TRABZON': 'TRTZX', 'YALOVA': 'TRYAL', 'YARIMCA': 'TRYAR', 'ZONGULDAK': 'TRZON', 'PORT-OF-SPAIN': 'TTPOS', 'POINT LISAS': 'TTPTS', 'SCARBOROUGH/TOBAGO': 'TTSCA', 'FUNAFUTI': 'TVFUN', 'KEELUNG': 'TWKEL', 'KAOHSIUNG': 'TWKHH', 'TAIPEI': 'TWTPE', 'TAICHUNG': 'TWTXG', 'DAR ES SALAAM': 'TZDAR', 'MTWARA': 'TZMYW', 'TANGA': 'TZTGT', 'ZANZIBAR': 'TZZNZ', 'CHORNOMORSK': 'UAILK', 'IZMAIL': 'UAIZM', 'ODESA': 'UAODS', 'RENI': 'UARNI', 'YALTA': 'UAYAL', 'YUZHNYY': 'UAYUZ', 'BALTIMORE': 'USBAL', 'BAR HARBOR': 'USBHB', 'BOSTON': 'USBOS', 'BROWNSVILLE': 'USBRO', 'CHARLESTON': 'USCHS', 'PT ANGELES': 'USCLM', 'CAPE CANAVERAL': 'USCPV', 'DUTCH HARBOR': 'USDUT', 'DAVISVILLE': 'USDVV', 'ELLISVILLE': 'USEVE', 'NEWARK': 'USEWR', 'KEY WEST': 'USEYW', 'FORT LAUDERDALE': 'USFLL', 'GULFPORT': 'USGPT', 'HONOLULU': 'USHNL', 'HOUSTON': 'USHOU', 'WILMINGTON, DE': 'USILG', 'WILMINGTON, NC': 'USILM', 'HILO': 'USITO', 'JACKSONVILLE': 'USJAX', 'KAWAIHAE': 'USKWH', 'LOS ANGELES': 'USLAX', 'LONG BEACH': 'USLGB', 'MIAMI': 'USMIA', 'MOBILE': 'USMOB', 'NEW ORLEANS': 'USMSY', 'NAWILIWILI': 'USNIJ', 'NEWPORT': 'USNPO', 'PORT HUENEME': 'USNTD', 'NEW YORK': 'USNYC', 'OAKLAND': 'USOAK', 'NORFOLK': 'USORF', 'PORTLAND, OR': 'USPDX', 'PORT EVERGLADES': 'USPEF', 'PHILADELPHIA': 'USPHL', 'PORTLAND, ME': 'USPWM', 'SAVANNAH': 'USSAV', 'SEATTLE': 'USSEA', 'SAN FRANCISCO': 'USSFO', 'TACOMA': 'USTIW', 'TAMPA': 'USTPA', 'UNALASKA': 'USUAA', 'MONTEVIDEO': 'UYMVD', 'NUEVA PALMIRA': 'UYNVP', 'PUNTA DEL ESTE': 'UYPDP', 'CAMPDEN PARK': 'VCCRP', 'KINGSTOWN, ST VINCENT': 'VCKTN', 'EL TABLAZO/MARACAIBO L': 'VEETV', 'GUARANAO BAY': 'VEGUB', 'LA GUAIRA': 'VELAG', 'PUERTO CABELLO': 'VEPBL', 'PUERTO LA CRUZ': 'VEPCZ', 'N. SOUND/VIRGIN GORDA': 'VGNSX', 'ROAD TOWN, TORTOLA': 'VGRAD', 'CHARLOTTE AMALIE, ST THOMAS': 'VICHA', 'SAINT THOMAS': 'VISTT', 'DA-NANG': 'VNDAD', 'HAIPHONG': 'VNHPH', 'NHA TRANG': 'VNNHA', 'PHUOC LONG': 'VNPHG', 'HO CHI MINH CITY': 'VNSGN', 'QUINHON': 'VNUIH', 'VUNG TAU': 'VNVUT', 'PORT VILA': 'VUVLI', 'APIA': 'WSAPW', 'ADEN': 'YEADE', 'HODEIDAH': 'YEHOD', 'MUKALLA': 'YEMKX', 'LONGONI': 'YTLON', 'CAPE TOWN': 'ZACPT', 'DURBAN': 'ZADUR', 'EAST LONDON': 'ZAELS', 'MOSSEL BAY': 'ZAMZY', 'PORT ELIZABETH': 'ZAPLZ', 'RICHARDS BAY': 'ZARCB', 'COEGA': 'ZAZBA'}

def convert_to_port_code(port_name):
    if isinstance(port_name, str):
        return PORT_CODE_MAP.get(port_name.strip().upper(), 'UNKNOWN')
    return 'UNKNOWN'


def merge_bl_columns(df):
    target_col = 'BL No.<Booking No.>'
    found = False
    if target_col in df.columns:
        found = True
    else:
        for possible in ['BL No.', 'BL.<Booking No.>', 'BL.No.<Booking No.>']:
            if possible in df.columns:
                df = df.rename(columns={possible: target_col})
                found = True
                break
    if found:
        df[target_col] = df[target_col].apply(
            lambda x: str(x).split('<')[0].strip() if pd.notnull(x) and '<' in str(x) else x
        )
    return df


def process_files(tsteam_bytes, itps_bytes, tsteam_filename,
                  pnit_vessel='', pnit_pod='', pnc_vessel='', pnc_pod=''):
    """
    메인 처리 로직 ─ IM_TS28.py 와 동일한 로직을 수행.
    """
    # ── 1. 파일 읽기 ─────────────────────────────────────────────────
    tsteam_data = pd.read_excel(io.BytesIO(tsteam_bytes), sheet_name=None)
    itps_data   = pd.read_excel(io.BytesIO(itps_bytes))

    # ── 2. BL 컬럼 통일 ─────────────────────────────────────────────
    itps_data = merge_bl_columns(itps_data)

    # ── 3. TSTEAM 시트 파싱 → ts_list ────────────────────────────────
    ts_list = []
    for sheet_name, sdf in tsteam_data.items():
        has_cntr = 'Cntr No.' in sdf.columns or 'Container Number' in sdf.columns
        if not has_cntr:
            continue
        cntr_col = 'Cntr No.' if 'Cntr No.' in sdf.columns else 'Container Number'
        for _, row in sdf.iterrows():
            por_value = row.get('POR', row.get('POL', ''))
            fpod_value = row.get('FPOD', row.get('FND', ''))
            ts_list.append({
                'Cntr No.': row[cntr_col],
                'POR':      por_value,
                'POL':      row.get('POL', ''),
                'POD':      row.get('POD', ''),
                'FPOD':     fpod_value,
                'REMARK':   row.get('REMARK', ''),
            })

    ts_df = pd.DataFrame(ts_list)

    # ── 4. ITPS 에 TS 데이터 매핑 ────────────────────────────────────
    if 'Equipment Number' in itps_data.columns and not ts_df.empty:
        def replace_values(row):
            match = ts_df[ts_df['Cntr No.'] == row['Equipment Number']]
            if not match.empty:
                row['Origin Load Port'] = match.iloc[0]['POR']
                row['Next POD']         = match.iloc[0]['POD']
                discharge_port = row.get('Discharge Port', '')
                if pd.notnull(discharge_port) and str(discharge_port).strip():
                    parts = str(discharge_port).split(',')
                    port_name = ','.join(parts[:-1]).strip() if len(parts) > 1 else str(parts[0]).strip()
                    code = convert_to_port_code(port_name)
                    if code != 'UNKNOWN':
                        row['Discharge Port'] = code
                row['Active'] = match.iloc[0]['REMARK']
            return row
        itps_data = itps_data.apply(replace_values, axis=1)

    # ── 5. VGM Weight 처리 & Weight 반올림 ──────────────────────────
    if 'Weight' in itps_data.columns:
        itps_data['Weight'] = itps_data['Weight'].apply(lambda x: round(float(x)) if pd.notnull(x) else 0)

    if 'VGM Weight' in itps_data.columns and 'Weight' in itps_data.columns:
        itps_data['VGM Weight'] = (
            itps_data['VGM Weight']
            .apply(lambda x: round(float(x)) if pd.notnull(x) else None)
            .fillna(itps_data['Weight'])
            .fillna(0)
            .astype(int)
        )

    if 'VGM Unit' in itps_data.columns:
        itps_data['VGM Unit'] = itps_data['VGM Unit'].fillna('KGS')
        def convert_lbs_to_kgs(row):
            if str(row.get('VGM Unit', '')).upper() == 'LBS':
                row['VGM Weight'] = round(float(row.get('VGM Weight', 0)) * 0.453592)
                row['VGM Unit']   = 'KGS (from LBS)'
            return row
        itps_data = itps_data.apply(convert_lbs_to_kgs, axis=1)

    # ── 6. Discharge Port – TS/KR 분기 ──────────────────────────────
    if 'Discharge Port' in itps_data.columns:
        def fill_discharge_port(row):
            dp = str(row.get('Discharge Port', '') or '')
            if dp and not dp.startswith('KR'):
                if pnit_vessel:
                    row['Loading Terminal']      = 'PNIT'
                    row['Loaded Vessel/Voyage']  = pnit_vessel
                    row['Next POD']              = pnit_pod
                    row['Discharge Port']        = pnit_pod
                elif pnc_vessel:
                    row['Loading Terminal']     = 'PNC'
                    row['Loaded Vessel/Voyage'] = pnc_vessel
                    row['Next POD']             = pnc_pod
                    row['Discharge Port']       = pnc_pod
            return row
        itps_data = itps_data.apply(fill_discharge_port, axis=1)

    # ── 7. L/T 분류 ──────────────────────────────────────────────────
    kr_ports = ['KRPUS', 'KRINC', 'KRKAN']
    def classify_pod(row):
        if str(row.get('Full/Empty', '')).upper() != 'E':
            dp = str(row.get('Discharge Port', '') or '')
            if dp in kr_ports:
                return dp
            else:
                return 'TS'
        return None

    if 'Full/Empty' in itps_data.columns and 'Discharge Port' in itps_data.columns:
        itps_data['L/T'] = itps_data.apply(classify_pod, axis=1)
    else:
        itps_data['L/T'] = 'TS'

    # ── 8. SUMMARY 생성 ──────────────────────────────────────────────
    if 'Type/Size' in itps_data.columns:
        summary_data = (
            itps_data.pivot_table(index='Type/Size', columns='L/T', aggfunc='size', fill_value=0)
            .reset_index()
        )
        
        # Ensure all KR ports and TS exist
        for port in ['KRPUS', 'KRKAN', 'KRINC', 'TS']:
            if port not in summary_data.columns:
                summary_data[port] = 0

        empty_counts = (
            itps_data[itps_data['Full/Empty'].astype(str).str.upper() == 'E']
            .groupby('Type/Size')
            .size()
            .reindex(summary_data['Type/Size'], fill_value=0)
            .values
        )
        summary_data['Empty'] = empty_counts
        
        # Calculate Total (Sum of KRPUS, KRKAN, KRINC, TS)
        summary_data['TOTAL'] = summary_data[['KRPUS', 'KRKAN', 'KRINC', 'TS']].sum(axis=1).astype(int)
        
        # Rename original Type/Size to TYPE/SIZE for consistency with user request
        summary_data = summary_data.rename(columns={'Type/Size': 'TYPE/SIZE'})
        
        for col in summary_data.columns:
            if summary_data[col].dtype != 'object':
                summary_data[col] = summary_data[col].astype(int)
        
        totals = summary_data.sum(numeric_only=True).astype(int)
        totals_row = pd.DataFrame([totals])
        totals_row.insert(0, 'TYPE/SIZE', 'Total')
        summary_data = pd.concat([summary_data, totals_row], ignore_index=True)
        
        # Final column order: KRPUS, KRKAN, KRINC, TS, TOTAL, TYPE/SIZE
        final_summary_cols = ['KRPUS', 'KRKAN', 'KRINC', 'TS', 'TOTAL', 'TYPE/SIZE']
        summary_data = summary_data[final_summary_cols]
    else:
        summary_data = pd.DataFrame({'TYPE/SIZE': ['No Data'], 'TOTAL': [0]})

    # ── 9. 터미널 이름 매핑 및 정렬 준비 ─────────────────────────────────
    terminal_map = {
        'PUSAN NEWPORT INTERNATIONAL TERMINAL-PNIT': 'PNIT',
        'PUSAN NEWPORT COMPANY LIMITED': 'PNC',
        'DONGWON GLOBAL TERMINAL (DGT)': 'DGT',
        'BUSAN CONTAINER TERMINAL - BCT': 'BCT',
        'HANJIN BUSAN NEW PORT CO., LTD': 'HJNC',
        'HMM PSA NEW-PORT TERMINAL (HPNT)': 'HPNT',
    }
    if 'Loading Terminal' in itps_data.columns:
        itps_data['Loading Terminal'] = itps_data['Loading Terminal'].replace(terminal_map)
        itps_data['Loading Terminal'] = itps_data['Loading Terminal'].fillna('')

    # ── 10. 정렬 (Loading Terminal 기준, 알파벳 순, 빈 값 하순) ──────────
    itps_data = itps_data.sort_values(
        by='Loading Terminal',
        key=lambda col: col.map(lambda x: (1, '') if (x is None or str(x).strip() == '') else (0, str(x).strip())),
        ascending=True
    )

    # ── 11. 컬럼 이름 변경 & 비고 처리 ───────────────────────────────────
    rename_map = {}
    if 'Next POD' in itps_data.columns: rename_map['Next POD'] = 'Stow Code'
    if 'Active' in itps_data.columns:   rename_map['Active'] = 'Remark'
    if rename_map:
        itps_data = itps_data.rename(columns=rename_map)

    if 'Remark' not in itps_data.columns: itps_data['Remark'] = ''
    itps_data['Remark'] = itps_data['Remark'].astype(str).replace(['nan', 'None'], '')

    if 'IMO Class' in itps_data.columns:
        itps_data['IMO Class'] = itps_data['IMO Class'].astype(str).replace(['nan', 'None'], '')

    def add_remarks(row):
        imo = str(row.get('IMO Class', '')).strip()
        reefer = row.get('Reefer Temp.', None)
        cur = str(row.get('Remark', '') or '').strip()
        if imo in ['2', '2.1', '2.2'] or (imo and pd.notna(reefer)):
            cur = (cur + ' 직반출 위험물').strip()
        if str(row.get('SOC', '')).upper() == 'Y':
            cur = (cur + ' SOC').strip()
        row['Remark'] = cur
        return row
    itps_data = itps_data.apply(add_remarks, axis=1)

    # ── 12. 터미널 축약 (나머지 컬럼) ───────────────────────────────────
    if 'Discharging Terminal' in itps_data.columns:
        itps_data['Discharging Terminal'] = itps_data['Discharging Terminal'].replace(terminal_map)

    # ── 13. 출력 컬럼 정렬 ──────────────────────────────────────────────
    columns_order = [
        'Equipment Number', 'BL No.<Booking No.>',
        'Discharge Vessel/Voyage', 'Discharging Terminal',
        'Loaded Vessel/Voyage', 'Loading Terminal',
        'Origin Load Port', 'Stow Code', 'Discharge Port',
        'Type/Size', 'Full/Empty', 'Weight', 'VGM Weight', 'VGM Unit',
        'Reefer or IMO', 'Reefer Temp.',
        'Dangerous Cargo', 'IMO Class', 'UN Number',
        'Vet. Control', 'Remark',
    ]
    itps_data = itps_data.reindex(columns=columns_order, fill_value='')

    # ── 14. 서브 탭 데이터 생성 ──────────────────────────────────────
    soc_cols = ['BL No.<Booking No.>', 'Equipment Number', 'Origin Load Port', 'Discharge Port', 'Full/Empty', 'Type/Size']
    soc_df = itps_data[itps_data['Remark'].astype(str).str.contains('SOC', na=False)][soc_cols].copy()
    soc_total = pd.DataFrame([['Total', '', '', '', '', len(soc_df)]], columns=soc_cols)

    dg_cols = ['BL No.<Booking No.>', 'Equipment Number', 'Full/Empty', 'Origin Load Port', 'Discharge Port', 'IMO Class', 'UN Number', 'Reefer Temp.']
    dg_df = itps_data[itps_data['Remark'].astype(str).str.contains('직반출 위험물', na=False)][dg_cols].copy()
    dg_total = pd.DataFrame([['Total', '', '', '', '', '', '', len(dg_df)]], columns=dg_cols)

    # ── 15. Excel 출력 ───────────────────────────────────────────────
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        itps_data.to_excel(writer, index=False, sheet_name='ITPS_DATA')
        summary_data.to_excel(writer, index=False, sheet_name='SUMMARY')
        soc_df.to_excel(writer, index=False, sheet_name='SOC')
        soc_total.to_excel(writer, index=False, header=False, startrow=len(soc_df) + 1, sheet_name='SOC')
        dg_df.to_excel(writer, index=False, sheet_name='직반출 위험물')
        dg_total.to_excel(writer, index=False, header=False, startrow=len(dg_df) + 1, sheet_name='직반출 위험물')

    output.seek(0)
    wb = openpyxl.load_workbook(output)
    ws_itps = wb['ITPS_DATA']

    fills = {
        'PNC':  PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid"),
        'PNIT': PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid"),
        'DGT':  PatternFill(start_color="E0FFE0", end_color="E0FFE0", fill_type="solid"),
        'BCT':  PatternFill(start_color="FFE4E1", end_color="FFE4E1", fill_type="solid"),
        'HJNC': PatternFill(start_color="F0E68C", end_color="F0E68C", fill_type="solid"),
        'HPNT': PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid"),
        'OLD PORT FEEDER TERMINAL': PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid"),
    }

    dis_col = load_col = None
    for cell in ws_itps[1]:
        if cell.value == 'Discharging Terminal': dis_col = cell.column_letter
        elif cell.value == 'Loading Terminal':  load_col = cell.column_letter

    def color_rows(sheet, col_let):
        for r in range(2, sheet.max_row + 1):
            val = str(sheet[f"{col_let}{r}"].value).strip()
            if val in fills:
                sheet[f"{col_let}{r}"].fill = fills[val]

    if dis_col: color_rows(ws_itps, dis_col)
    if load_col: color_rows(ws_itps, load_col)

    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            for cell in col:
                try: 
                    if cell.value:
                        l = len(str(cell.value).encode('utf-8'))
                        if l > max_len: max_len = l
                except: pass
            ws.column_dimensions[col[0].column_letter].width = min(max_len * 0.9 + 2, 50)

    # ── 16. 요약 데이터 생성 (UI 용) ───────────────────────────────────
    summary_list = summary_data.to_dict(orient='records')
    soc_list = soc_df.to_dict(orient='records')
    dg_list = dg_df.to_dict(orient='records')

    final_io = io.BytesIO()
    wb.save(final_io)
    final_io.seek(0)
    
    return {
        'excel': final_io,
        'summary_data': summary_list,
        'soc_data': soc_list,
        'dg_data': dg_list
    }

# ══════════════════════════════════════════════════════════════════════
# FLASK
# ══════════════════════════════════════════════════════════════════════
app = Flask(__name__)

@app.route('/')
def index():
    return open('index.html', encoding='utf-8').read()

@app.route('/process', methods=['POST'])
def process():
    try:
        import base64
        t_file = request.files['tsteam']
        i_file = request.files['itps']
        
        result = process_files(t_file.read(), i_file.read(), t_file.filename,
                              request.form.get('pnit_vessel',''), request.form.get('pnit_pod',''),
                              request.form.get('pnc_vessel',''), request.form.get('pnc_pod',''))
        
        excel_io = result['excel']
        b64_excel = base64.b64encode(excel_io.getvalue()).decode('utf-8')
        
        base = re.sub(r'^최종파일\s*-\s*', '', os.path.splitext(t_file.filename)[0]).strip()
        filename = f"{base} merged file.xlsx"
        
        return jsonify({
            'filename': filename,
            'file_data': b64_excel,
            'summary': result['summary_data'],
            'soc': result['soc_data'],
            'dg': result['dg_data']
        })
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5050)
