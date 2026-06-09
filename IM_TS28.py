import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import filedialog
from tkinter import ttk

import pandas as pd
import os

from tabulate import tabulate
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl import Workbook

from datetime import datetime, timedelta
import pyperclip

#pyinstaller -w -F --add-binary="C:/Users/kod03/AppData/Local/Programs/Python/Python311/tcl/tkdnd2.8;tkdnd2.8" example.py

port_svc = {
    ('PNIT', 'CNDLC', ''): 'PERTIWI',
    ('PNIT', 'CNNGB', ''): 'DRAGON',
    ('PNIT', 'CNNSA', ''): 'DOLPHIN',
    ('PNIT', 'CNSHA', ''): 'DRAGON',
    ('PNIT', 'CNSHK', ''): 'NEW FALCON SERVICE',
    ('PNIT', 'CNTAO', ''): 'NEW FALCON SERVICE',
    ('PNIT', 'CNTXG', ''): 'PERTIWI',
    ('PNIT', 'CNXMN', ''): 'NEW FALCON SERVICE',
    ('PNIT', 'CNYTN', ''): 'DRAGON',
    ('PNIT', 'HKHKG', ''): 'ANDES EXPRESS SERVICES',
    ('PNIT', 'JPHKT', ''): 'KAGUYA',
    ('PNIT', 'JPNGO', 'R'): 'ORIGAMI',
    ('PNIT', 'JPOSA', ''): 'KAGUYA',
    ('PNIT', 'JPTYO', 'R'): 'ORIGAMI',
    ('PNIT', 'JPUKB', ''): 'KAGUYA',
    ('PNIT', 'JPYKK', 'R'): 'ORIGAMI',
    ('PNIT', 'JPYOK', ''): 'AZTEC',
    ('PNIT', 'KRINC', ''): 'PERTIWI',
    ('PNIT', 'RUVVO', ''): 'SUNRISE',
    ('PNIT', 'SGSIN', ''): 'NEW FALCON SERVICE',
    ('PNIT', 'VNVUT', 'A'): 'ORIGAMI',
    ('PNC', 'CNDLC', ''): 'PERTIWI',
    ('PNC', 'CNNGB', ''): 'AFRICA EXPRESS',
    ('PNC', 'CNSHA', ''): 'JADE',
    ('PNC', 'CNNSA', ''): 'JADE',
    ('PNC', 'CNYTN', ''): 'JADE',
    ('PNC', 'CNTAO', ''): 'NEW FALCON SERVICE',
    ('PNC', 'CNTXG', ''): 'PERTIWI',
    ('PNC', 'JPYKK', 'R'): 'ORIGAMI',
    ('PNC', 'CNSHK', ''): 'AFRICA EXPRESS EXPRESS',
    ('PNC', 'CNXMN', ''): 'JADE',
    ('PNC', 'JPTYO', 'R'): 'ORIGAMI',
    ('PNC', 'HKHKG', ''): 'ANDES EXPRESS SERVICES',
    ('PNC', 'JPHKT', ''): 'KAGUYA',
    ('PNC', 'JPNGO', 'R'): 'ORIGAMI',
    ('PNC', 'JPOSA', ''): 'KAGUYA',
    ('PNC', 'JPUKB', ''): 'KAGUYA',
    ('PNC', 'JPYOK', ''): 'MAPLE',
    ('PNC', 'KRINC', ''): 'PERTIWI',
    ('PNC', 'RUVVO', ''): 'SUNRISE',
    ('PNC', 'SGSIN', ''): 'JADE',
    ('PNC', 'VNVUT', 'A'): 'ORIGAMI'
}

def convert_to_port_code(port_name):
    port_codes = {'AJMAN': 'AEAJM', 'ABU DHABI': 'AEAUH', 'DUBAI': 'AEDXB', 'AL - FUJAYRAH': 'AEFJR', 'JEBEL ALI': 'AEJEA', 'KHOR AL FAKKAN': 'AEKLF', 'PORT RASHID': 'AEPRA', 'UMM AL QAIWAIN': 'AEQIW', 'RAS AL KHAIMAH': 'AERKT', 'SHARJAH': 'AESHJ', 'YAS ISLAND': 'AEYAS', 'ANTIGUA': 'AGANU', 'ROAD BAY': 'AIRBY', 'DURRES': 'ALDRZ', 'SARANDE': 'ALSAR', 'LUANDA': 'AOLAD', 'LOBITO': 'AOLOB', 'NAMIBE': 'AOMSZ', 'BUENOS AIRES': 'ARBUE', 'CAMPANA': 'ARCMP', 'LA PLATA': 'ARLPG', 'PUERTO MADRYN': 'ARPMY', 'ROSARIO': 'ARROS', 'USHUAIA': 'ARUSH', 'ZARATE': 'ARZAE', 'PAGO PAGO': 'ASPPG', 'ABBOT POINT': 'AUABP', 'ADELAIDE': 'AUADL', 'ALBANY': 'AUALH', 'BELL BAY': 'AUBEL', 'BRISBANE': 'AUBNE', 'BOOBY ISLAND': 'AUBOO', 'CAIRNS': 'AUCNS', 'DARWIN': 'AUDRW', 'ESPERANCE': 'AUEPR', 'FREMANTLE': 'AUFRE', 'GLADSTONE': 'AUGLT', 'HOBART': 'AUHBA', 'HAY POINT': 'AUHPT', 'MELBOURNE': 'AUMEL', 'NEWCASTLE': 'GBNCS', 'PORT HEDLAND': 'AUPHE', 'PORT KEMBLA': 'AUPKL', 'SYDNEY': 'CASYD', 'ORANJESTAD': 'AWORJ', 'BRIDGETOWN': 'BBBGI', 'CHATTOGRAM': 'BDCGP', 'MONGLA': 'BDMGL', 'ANTWERP': 'BEANR', 'GENT (GHENT)': 'BEGNE', 'ZEEBRUGGE': 'BEZEE', 'BURGAS': 'BGBOJ', 'VARNA': 'BGVAR', 'BAHRAIN': 'BHKBS', 'COTONOU': 'BJCOO', 'HAMILTON': 'BMBDA', 'KINGS WHARF': 'BMKWF', 'SINT EUSTATIUS': 'BQEUX', 'ARRAIAL DO CABO': 'BRACB', 'ANGRA DOS REIS': 'BRADR', 'ARTUR NOGUEIRA': 'BRANG', 'BELEM': 'BRBEL', 'BUZIOS': 'BRBZC', 'CAMBORIU': 'BRCBU', 'CABEDELO': 'BRCDO', 'CARMO DO PARANAIBA': 'BRCOP', 'FORTALEZA': 'BRFOR', 'IMBITUBA': 'BRIBB', 'ILHABELA': 'BRIBE', 'ILHA GRANDE': 'BRIGE', 'ITAGUAI': 'BRIGI', 'ITAPOA': 'BRIOA', 'ILHEUS': 'BRIOS', 'ITAJAI': 'BRITJ', 'MANAUS': 'BRMAO', 'MACEIO': 'BRMCZ', 'NAVEGANTES': 'BRNVT', 'PORTO BELO': 'BRPBO', 'PECEM': 'BRPEC', 'PARANAGUA': 'BRPNG', 'PORTO VELHO': 'BRPVH', 'CABO FRIO': 'BRQCK', 'RECIFE': 'BRREC', 'RIO GRANDE': 'BRRIG', 'RIO DE JANEIRO': 'BRRIO', 'SAO FRANCISCO DO SUL': 'BRSFS', 'SALVADOR': 'BRSSA', 'SANTOS': 'BRSSZ', 'SANTAREM': 'BRSTM', 'SUAPE': 'BRSUA', 'UBATUBA': 'BRUBT', 'VITORIA': 'BRVIX', 'VILA DO CONDE': 'BRVLC', 'COCO CAY': 'BSCOC', 'FREEPORT, GRAND BAHAMA': 'BSFPO', 'GREAT STIRRUP CAY': 'BSGSC', 'LITTLE SAN SALVADOR': 'BSHMC', 'LUCAYA': 'BSLUC', 'NASSAU': 'BSNAS', 'BELIZE CITY': 'BZBZE', 'BAIE COMEAU': 'CABCO', 'BECANCOUR': 'CABEC', 'CORNER BROOK': 'CACBK', 'CHARLOTTETOWN': 'CACHA', 'GASPE': 'CAGPE', 'HALIFAX': 'CAHAL', 'HAVRE-SAINT-PIERRE': 'CAHSP', 'LA BAIE': 'CALBA', 'MONTREAL': 'CAMTR', 'PRINCE RUPERT': 'CAPRR', 'QUEBEC': 'CAQUE', 'SAINT JOHN': 'CASJB', 'TROIS-RIVIERES (THREE RIVERS)': 'CATRR', 'VANCOUVER': 'CAVAN', 'COCOS ISLANDS': 'CCCCK', 'MATADI': 'CDMAT', 'POINTE NOIRE': 'CGPNR', 'ABIDJAN': 'CIABJ', 'SAN-PEDRO': 'CISPY', 'AITUTAKI': 'CKAIT', 'RAROTONGA': 'CKRAR', 'ANTOFAGASTA': 'CLANF', 'ARICA': 'CLARI', 'CORONEL': 'CLCNL', 'COQUIMBO': 'CLCQQ', 'ISLA DE PASCUA': 'CLIPC', 'IQUIQUE': 'CLIQQ', 'LIRQUEN': 'CLLQN', 'PUERTO ANGAMOS': 'CLPAG', 'MAGELLAN STRAIT': 'CLPPY', 'SAN ANTONIO': 'CLSAI', 'SAN VICENTE': 'CLSVE', 'VALPARAISO': 'CLVAP', 'DOUALA': 'CMDLA', 'TIESHAN': 'CNBHT', 'GUANGZHOU': 'CNCAN', 'CAOFEIDIAN': 'CNCFD', 'DA CHAN BAY': 'CNDCB', 'DALIAN': 'CNDLC', 'FANGCHENG': 'CNFAN', 'FUZHOU': 'CNFOC', 'HAIKOU': 'CNHAK', 'HUANGHUA': 'CNHNH', 'HUANGPU': 'CNHUA', 'JINGTANG (TANGSHAN)': 'CNJIN', 'LU-HUA SHAN': 'CNLUH', 'LIANYUNGANG': 'CNLYG', 'MAWEI': 'CNMAW', 'MAWAN': 'CNMWN', 'NINGBO': 'CNNGB', 'NANJING': 'CNNKG', 'NANSHA': 'CNNSA', 'NANTONG': 'CNNTG', 'QINZHOU': 'CNQZH', 'SHANGHAI': 'CNSHA', 'SANSHAN': 'CNSHG', 'SHEKOU': 'CNSHK', 'QINHUANGDAO': 'CNSHP', 'SHANTOU': 'CNSWA', 'TAICANG': 'CNTAG', 'QINGDAO': 'CNTAO', 'TIANJINXINGANG': 'CNTXG', 'WEIHAI': 'CNWEI', 'XIAMEN': 'CNXMN', 'XIANGSHUI': 'CNXSI', 'YANTAI': 'CNYNT', 'YANTIAN': 'CNYTN', 'ZHANGJIAGANG': 'CNZJG', 'ZHOUSHAN': 'CNZOS', 'BARRANQUILLA': 'COBAQ', 'BUENAVENTURA': 'COBUN', 'CARTAGENA': 'ESCAR', 'TURBO': 'COTRB', 'CALDERA': 'CRCAL', 'PUERTO LIMON': 'CRLIO', 'MOIN': 'CRMOB', 'NUEVA GERONA': 'CUGER', 'LA HABANA': 'CUHAV', 'MARIEL': 'CUMAR', 'MINDELO': 'CVMIN', 'PRAIA': 'CVRAI', 'SAL REI': 'CVSAR', 'CURACAO': 'CWCUR', 'FAMAGUSTA': 'CYFMG', 'LIMASSOL': 'CYLMS', 'BRUNSBUTTEL': 'DEBRB', 'BREMEN': 'DEBRE', 'BREMERHAVEN': 'DEBRV', 'ELSFLETH': 'DEELS', 'EMDEN': 'DEEME', 'HAMBURG': 'DEHAM', 'HELGOLAND': 'DEHGL', 'KIEL': 'DEKEL', 'LUBECK': 'DELBC', 'WARNEMUNDE': 'DEWAR', 'WILHELMSHAVEN': 'DEWVN', 'DJIBOUTI': 'DJJIB', 'AALBORG': 'DKAAL', 'AARHUS': 'DKAAR', 'COPENHAGEN': 'DKCPH', 'FREDERICIA': 'DKFRC', 'GREAT BELT': 'DKGBT', 'HVIDE SANDE': 'DKHVS', 'KALUNDBORG': 'DKKAL', 'KOLIND': 'DKKLD', 'KERTEMINDE': 'DKKTD', 'ODENSE': 'DKODE', 'RONNE': 'DKRNN', 'SKAGEN': 'DKSKA', 'PORTSMOUTH': 'USPTM', 'ROSEAU': 'DMRSU', 'CAUCEDO': 'DOCAU', 'RIO HAINA': 'DOHAI', 'PUERTO PLATA': 'DOPOP', 'ANNABA': 'DZAAE', 'ALGER': 'DZALG', 'ARZEW': 'DZAZW', 'BEJAIA': 'DZBJA', 'DJEN-DJEN': 'DZDJE', 'GHAZAOUET': 'DZGHZ', 'ORAN': 'DZORN', 'SKIKDA': 'DZSKI', 'ESMERALDAS': 'ECESM', 'GUAYAQUIL': 'ECGYE', 'LA LIBERTAD': 'ECLLD', 'PUERTO BOLIVAR': 'ECPBO', 'POSORJA': 'ECPSJ', 'TALLINN': 'EETLL', 'ALEXANDRIA OLD PORT': 'EGALY', 'DAMIETTA': 'EGDAM', 'ALEXANDRIA EL DEKHEILA': 'EGEDK', 'PORT SAID EAST': 'EGPSE', 'PORT SAID WEST': 'EGPSW', 'SAFAGA': 'EGSGA', 'SOKHNA PORT': 'EGSOK', 'SHARM ASH SHAYKH': 'EGSSH', 'SUEZ': 'EGSUZ', 'ASSAB': 'ERASA', 'ARRECIFE DE LANZAROTE': 'ESACE', 'MALAGA': 'ESAGP', 'ALICANTE': 'ESALC', 'ALGECIRAS': 'ESALG', 'ALMAGRO': 'ESALM', 'BARCELONA': 'ESBCN', 'BILBAO': 'ESBIO', 'CADIZ': 'ESCAD', 'CASTELLON DE LA PLANA': 'ESCAS', 'CEUTA': 'ESCEU', 'FERROL': 'ESFRO', 'PUERTO DEL ROSARIO-FUERTEVENTURA': 'ESFUE', 'GIJON': 'ESGIJ', 'GRANJA DE SAN IDELFONSO': 'ESGJI', 'HUELVA': 'ESHUV', 'IBIZA': 'ESIBZ', 'LA CORUNA': 'ESLCG', 'ALMERIA': 'ESLEI', 'LAS PALMAS': 'ESLPA', 'MAHON, MENORCA': 'ESMAH', 'MARIN, PONTEVEDRA': 'ESMPG', 'PALAMOS': 'ESPAL', 'PALMA DE MALLORCA': 'ESPMI', 'ROSAS': 'ESROS', 'SAGUNTO': 'ESSAG', 'SANTA CRUZ DE TENERIFE': 'ESSCT', 'SANTA CRUZ DE LA PALMA': 'ESSPC', 'SAN SEBASTIAN DE LA GOMERA': 'ESSSG', 'SEVILLA': 'ESSVQ', 'TARRAGONA': 'ESTAR', 'VIGO': 'ESVGO', 'VALENCIA': 'ESVLC', 'HELSINKI': 'FIHEL', 'KEMI': 'FIKEM', 'KOKKOLA (KARLEBY)': 'FIKOK', 'KOTKA': 'FIKTK', 'OULU (ULEABORG)': 'FIOUL', 'RAUMA': 'FIRAU', 'TORNIO (TORNEA)': 'FITOR', 'LAUTOKA': 'FJLTK', 'SUVA': 'FJSUV', 'THORSHAVN': 'FOTHO', 'AJACCIO': 'FRAJA', 'BREST': 'FRBES', 'BORDEAUX': 'FRBOD', 'CANNES': 'FRCEQ', 'CHERBOURG': 'FRCER', 'DUNKERQUE': 'FRDKK', 'FOS-SUR-MER': 'FRFOS', 'GENNEVILLIERS': 'FRGVL', 'HONFLEUR': 'FRHON', 'LE HAVRE': 'FRLEH', 'LA ROCHELLE': 'FRLRH', 'LE VERDON': 'FRLVE', 'MARSEILLE': 'FRMRS', 'MONTOIR-DE-BRETAGNE': 'FRMTX', 'NICE': 'FRNCE', 'SETE': 'FRSET', 'SIX-FOURS-LES-PLAGES': 'FRSFP', 'ST NAZAIRE': 'FRSNR', 'ST MARCEL': 'FRSTM', 'ST TROPEZ': 'FRSTP', 'TOULON': 'FRTLN', 'ROUEN': 'FRURO', 'VILLEFRANCHE-SUR-MER': 'FRVFM', 'LIBREVILLE': 'GALBV', 'PORT GENTIL': 'GAPOG', 'BELFAST': 'GBBEL', 'BRISTOL': 'GBBRS', 'DOVER': 'GBDVR', 'FALMOUTH': 'JMFMH', 'FELIXSTOWE': 'GBFXT', 'GRANGEMOUTH': 'GBGRG', 'GREENOCK': 'GBGRK', 'HARWICH': 'GBHRW', 'HULL': 'GBHUL', 'IMMINGHAM': 'GBIMM', 'LERWICK': 'GBLER', 'LONDON GATEWAY PORT': 'GBLGP', 'LIVERPOOL': 'GBLIV', 'PORTREE': 'GBPRT', 'PORTBURY': 'GBPRU', 'PORTLAND': 'GBPTL', 'SOUTH QUEENSFERRY': 'GBSOQ', 'SOUTHAMPTON': 'GBSOU', 'SOUTH SHIELDS': 'GBSSH', 'TEESPORT': 'GBTEE', 'THAMESPORT': 'GBTHP', 'TILBURY': 'GBTIL', 'GRENADA': 'GDGND', "SAINT GEORGE'S": 'GDSTG', 'BATUMI': 'GEBUS', 'POTI': 'GEPTI', 'ST PETER PORT': 'GGSPT', 'TEMA': 'GHTEM', 'TAKORADI': 'GHTKD', 'GIBRALTAR': 'GIGIB', 'NUUK (GODTHAAB)': 'GLGOH', 'PAAMIUT (FREDRIKSHAAB)': 'GLJFR', 'BANJUL': 'GMBJL', 'CONAKRY': 'GNCKY', 'POINTE-A-PITRE': 'GPPTP', 'ARGOSTOLION': 'GRARM', 'KERKIRA (CORFU)': 'GRCFU', 'CANEA (CHANIA)': 'GRCHQ', 'ELEFSIS (ELEVSIS)': 'GRELE', 'GYTHION': 'GRGYT', 'HERAKLION': 'GRHER', 'MYKONOS': 'GRJMK', 'SYROS (SYRA)': 'GRJSY', 'THIRA': 'GRJTR', 'KATAKOLON': 'GRKAK', 'KOS': 'GRKGS', 'KALILIMENES': 'GRKLL', 'KALAMATA': 'GRKLX', 'LAURIUM (LAVRION)': 'GRLAV', 'MOUDHROS': 'GRMDR', 'MONEMVASIA': 'GRMON', 'NAFPLION': 'GRNAF', 'PIRAEUS': 'GRPIR', 'PATMOS': 'GRPMS', 'RHODES': 'GRRHO', 'SOUDA': 'GRSDH', 'SKARAMANGAS': 'GRSKA', 'THESSALONIKI': 'GRSKG', 'TILOS': 'GRTIL', 'VOLOS': 'GRVOL', 'ZAKYNTHOS': 'GRZTH', 'PUERTO BARRIOS': 'GTPBR', 'PUERTO QUETZAL': 'GTPRQ', 'PUERTO SANTO TOMAS DE CASTILLA': 'GTSTC', 'BISSAU': 'GWOXB', 'GEORGETOWN': 'GYGEO', 'HONG KONG': 'HKHKG', 'PUERTO CASTILLA': 'HNPCA', 'PUERTO CORTES': 'HNPCR', 'ROATAN': 'HNRTB', 'SAN LORENZO': 'HNSLO', 'DUBROVNIK': 'HRDBV', 'PLOCE': 'HRPLE', 'RIJEKA': 'HRRJK', 'SPLIT': 'HRSPU', 'ZADAR': 'HRZAD', 'GONAIVES': 'HTGVS', 'LABADIE': 'HTLAB', 'PORT AU PRINCE': 'HTPAP', 'BELAWAN, SUMATRA': 'IDBLW', 'BENOA, BALI': 'IDBOA', 'BALIKPAPAN, KALIMANTAN': 'IDBPN', 'BATAM ISLAND': 'IDBTM', 'JAMBI, SUMATRA': 'IDDJB', 'JAKARTA, JAVA': 'IDJKT', 'MAKASSAR': 'IDMAK', 'MANGOLE': 'IDMAL', 'PADANG': 'IDPDG', 'PERAWANG': 'IDPER', 'PALEMBANG, SUMATRA': 'IDPLM', 'PANJANG': 'IDPNJ', 'PONTIANAK, KALIMANTAN': 'IDPNK', 'PERAWANG, SUMATRA': 'IDPWG', 'SEMARANG': 'IDSRG', 'SURABAYA': 'IDSUB', 'TABONEO': 'IDTAB', 'TANJUNG BARA, KL': 'IDTBA', 'UJUNG PANDANG, SULAWESI': 'IDUPG', 'DUN LAOGHAIRE': 'IEDLG', 'DUBLIN': 'IEDUB', 'GREENCASTLE': 'IEGRE', 'CORK': 'IEORK', 'WATERFORD': 'IEWAT', 'ASHDOD': 'ILASH', 'ELAT (EILATH)': 'ILETH', 'HAIFA': 'ILHFA', 'ALANG SBY': 'INALA', 'BHAVNAGAR': 'INBHU', 'MUMBAI': 'INBOM', 'KOLKATA': 'INCCU', 'COCHIN': 'INCOK', 'ENNORE': 'INENR', 'HALDIA': 'INHAL', 'HAZIRA PORT/SURAT': 'INHZA', 'KANDLA': 'INIXY', 'KAKINADA': 'INKAK', 'KATTUPALLI': 'INKAT', 'KRISHNAPATNAM': 'INKRI', 'CHENNAI': 'INMAA', 'MARMUGAO (MARMAGAO)': 'INMRM', 'MUNDRA': 'INMUN', 'NEW MANGALORE': 'INNML', 'NHAVA SHEVA': 'INNSA', 'PIPAVAV (VICTOR) PORT': 'INPAV', 'PARADIP GARH': 'INPRT', 'TUTICORIN': 'INTUT', 'VISAKHAPATNAM': 'INVTZ', 'UMM QASR PT': 'IQUQR', 'AKUREYRI': 'ISAKU', 'ISAFJORDUR - HOFN': 'ISISA', 'REYKJAVIK': 'ISREY', 'ALGHERO': 'ITAHO', 'ANCONA': 'ITAOI', 'AUGUSTA PORT OF CATANIA': 'ITAUG', 'BRINDISI': 'ITBDS', 'BARI': 'ITBRI', 'CAGLIARI': 'ITCAG', 'CATANIA': 'ITCTA', 'CIVITAVECCHIA': 'ITCVV', 'GIOIA TAURO': 'ITGIT', 'GENOA': 'ITGOA', 'ISOLA SANTO STEFANO': 'ITISS', 'LEGHORN': 'ITLIV', 'MONFALCONE': 'ITMNF', 'MESSINA': 'ITMSN', 'NAPLES': 'ITNAP', 'OLBIA': 'ITOLB', 'PALAZZOLO DELLO STELLA': 'ITPAL', 'MARGHERA': 'ITPMA', 'PALERMO': 'ITPMO', 'PORTOPALO': 'ITPPL', 'PORTOFINO': 'ITPTF', 'POZZALLO': 'ITPZL', 'SASSARI': 'ITQSS', 'RAVENNA': 'ITRAN', 'SORRENTO': 'ITRRO', 'SALERNO': 'ITSAL', 'SIRACUSA': 'ITSIR', 'LA SPEZIA': 'ITSPE', 'SAVONA': 'ITSVN', 'TARANTO': 'ITTAR', 'TRAPANI': 'ITTPS', 'TRIESTE': 'ITTRS', 'VENICE': 'ITVCE', 'VADO LIGURE': 'ITVDL', 'KINGSTON': 'JMKIN', 'MONTEGO BAY': 'JMMBJ', 'OCHO RIOS': 'JMOCJ', "AL 'AQABAH": 'JOAQJ', 'ABURATSU': 'JPABU', 'AOMORI': 'JPAOJ', 'AKITA': 'JPAXT', 'BEPPU, SHIMANE': 'JPBEP', 'CHIBA': 'JPCHB', 'FUKUYAMA, HIROSHIMA': 'JPFKY', 'FUKUOKA': 'JPFUK', 'HIBIKISHINKO': 'JPHBK', 'HACHINOHE, AOMORI': 'JPHHE', 'HITACHINAKA': 'JPHIC', 'HIROSHIMA': 'JPHIJ', 'HIMEJI': 'JPHIM', 'HAKODATE': 'JPHKD', 'HAKATA, FUKUOKA': 'JPHKT', 'HAMADA': 'JPHMD', 'HOSOSHIMA': 'JPHSM', 'IMABARI': 'JPIMB', 'IMARI': 'JPIMI', 'ISHIKARI': 'JPISI', 'IWAKUNI': 'JPIWK', 'IYOMISHIMA': 'JPIYM', 'KOCHI': 'JPKCZ', 'NIIGATA': 'JPKIJ', 'KAMAISHI': 'JPKIS', 'KUMAMOTO': 'JPKMJ', 'KANAZAWA': 'JPKNZ', 'KAGOSHIMA': 'JPKOJ', 'KURE, HIROSHIMA': 'JPKRE', 'KASHIMA, IBARAKI': 'JPKSM', 'KUSHIRO': 'JPKUH', 'KAWASAKI': 'JPKWS', 'MAIZURU': 'JPMAI', 'MIIKE, FUKUOKA': 'JPMII', 'MIZUSHIMA': 'JPMIZ', 'MOJI/KITAKYUSHU': 'JPMOJ', 'MURORAN': 'JPMUR', 'MATSUYAMA': 'JPMYJ', 'MIYAKO, IWATE': 'JPMYK', 'NAHA, OKINAWA': 'JPNAH', 'NAKANOSEKI': 'JPNAN', 'NAOETSU': 'JPNAO', 'NAGOYA': 'JPNGO', 'NAGASAKI': 'JPNGS', 'OHFUNATO': 'JPOFT', 'OITA': 'JPOIT', 'OMAEZAKI': 'JPOMZ', 'ONAHAMA': 'JPONA', 'OSAKA': 'JPOSA', 'OTAKE': 'JPOTK', 'SHIBUSHI': 'JPSBS', 'SENDAI, MIYAGI': 'JPSDJ', 'SHIMONOSEKI': 'JPSHS', 'SAKATA': 'JPSKT', 'SAKAIMINATO': 'JPSMN', 'SHIMIZU': 'JPSMZ', 'SATSUMASENDAI': 'JPSTS', 'TAKAMATSU': 'JPTAK', 'TOYOHASHI': 'JPTHS', 'TOKUSHIMA': 'JPTKS', 'TOKUYAMA': 'JPTKY', 'TOMAKOMAI': 'JPTMK', 'TOYAMASHINKO': 'JPTOS', 'TSURUGA': 'JPTRG', 'TOKYO': 'JPTYO', 'UBE': 'JPUBJ', 'KOBE': 'JPUKB', 'WAKAYAMA': 'JPWAK', 'YATSUSHIRO': 'JPYAT', 'YOKKAICHI': 'JPYKK', 'YOKOHAMA': 'JPYOK', 'EMBAKASI': 'KEEMB', 'MOMBASA': 'KEMBA', 'KAMPONG SAOM (SIHANOUKVILLE)': 'KHKOS', 'PHNOM PENH': 'KHPNH', 'MUTSAMUDU': 'KMMUT', 'MORONI': 'KMYVA', 'BASSETERRE, ST KITTS': 'KNBAS', 'NEVIS': 'KNNEV', 'GOSEONG-GUN': 'KRGSO', 'INCHEON': 'KRINC', 'GANGNEUNG': 'KRKAG', 'GWANGYANG': 'KRKAN', 'POHANG': 'KRKPO', 'MOKPO': 'KRMOK', 'OKPO/GEOJE': 'KROKP', 'BUSAN': 'KRPUS', 'SAMCHEONPO/SACHEON': 'KRSCP', 'SEOGWIPO': 'KRSPO', 'DANGJIN': 'KRTJI', 'TONGYEONG': 'KRTYG', 'ULSAN': 'KRUSN', 'YEOSU': 'KRYOS', 'SHUAIBA': 'KWSAA', 'SHUWAIKH': 'KWSWK', 'GRAND CAYMAN': 'KYGCM', 'ALMATY': 'KZALA', 'BEIRUT': 'LBBEY', 'CASTRIES': 'LCCAS', 'ST LUCIA APT': 'LCSLU', 'COLOMBO': 'LKCMB', 'GALLE': 'LKGAL', 'HAMBANTOTA': 'LKHBA', 'TRINCOMALEE': 'LKTRR', 'MONROVIA': 'LRMLW', 'KLAIPEDA': 'LTKLJ', 'RIGA': 'LVRIX', 'BINGAZI': 'LYBEN', 'KHOMS': 'LYKHO', 'MISURATA': 'LYMRA', 'TRIPOLI': 'LYTIP', 'AGADIR': 'MAAGA', 'CASABLANCA': 'MACAS', 'NADOR': 'MANDR', 'TANGER MED': 'MAPTM', 'MONTE-CARLO': 'MCMCM', 'GIURGIULESTI': 'MDGIU', 'BAR': 'MEBAR', 'BIJELA': 'MEBIJ', 'KOTOR': 'MEKOT', 'DIEGO SUAREZ': 'MGDIE', 'EHOALA (TOLAGNARO)': 'MGEHL', 'FORT DAUPHIN (TOALAGNARO)': 'MGFTU', 'MAJUNGA': 'MGMJN', 'NOSY-BE': 'MGNOS', 'SAINTE MARIE': 'MGSMS', 'TULEAR': 'MGTLE', 'TAMATAVE': 'MGTMM', 'VOHEMAR': 'MGVOH', 'MANAKARA': 'MGWVK', 'YANGON': 'MMRGN', 'MACAU': 'MOMFM', 'FORT-DE-FRANCE': 'MQFDF', 'NOUADHIBOU': 'MRNDB', 'NOUAKCHOTT': 'MRNKC', 'LITTLE BAY': 'MSLTB', 'PLYMOUTH': 'MSPLY', 'MARSAXLOKK': 'MTMAR', 'VALLETTA': 'MTMLA', 'PORT LOUIS': 'MUPLU', 'MALE': 'MVMLE', 'ALTAMIRA': 'MXATM', 'COSTA MAYA': 'MXCOM', 'COZUMEL': 'MXCZM', 'ENSENADA': 'MXESE', 'GUAYMAS': 'MXGYM', 'LAZARO CARDENAS': 'MXLZC', 'MAZATLAN': 'MXMZT', 'PROGRESO': 'MXPGO', 'PUERTO MORELOS': 'MXPMS', 'TAMPICO': 'MXTAM', 'TULUM': 'MXTUY', 'VERACRUZ': 'MXVER', 'MANZANILLO': 'PAMIT', 'KOTA KINABALU, SABAH': 'MYBKI', 'BINTULU, SARAWAK': 'MYBTU', 'KUCHING, SARAWAK': 'MYKCH', 'KUANTAN': 'MYKUA', 'LABUAN, SABAH': 'MYLBU', 'LANGKAWI': 'MYLGK', 'MALACCA': 'MYMKZ', 'MIRI, SARAWAK': 'MYMYY', 'PENANG': 'MYPEN', 'PASIR GUDANG, JOHOR': 'MYPGU', 'PORT KLANG (PELABUHAN KLANG)': 'MYPKG', 'SIBU, SARAWAK': 'MYSBW', 'SANDAKAN, SABAH': 'MYSDK', 'TANJUNG PELEPAS': 'MYTPP', 'TAWAU, SABAH': 'MYTWU', 'BEIRA': 'MZBEW', 'BAZARUTO ISLAND': 'MZBZB', 'NACALA': 'MZMNC', 'MAPUTO': 'MZMPM', 'MASSINGA': 'MZMSG', 'PEMBA': 'MZPOL', 'QUELIMANE': 'MZUEL', 'LUDERITZ': 'NALUD', 'WALVIS BAY': 'NAWVB', 'NOUMEA': 'NCNOU', 'VAVOUTO': 'NCVAV', 'APAPA': 'NGAPP', 'LEKKI': 'NGLKK', 'ONNE': 'NGONN', 'PORT HARCOURT': 'NGPHC', 'TINCAN/LAGOS': 'NGTIN', 'CORINTO': 'NICIO', 'MANAGUA': 'NIMGA', 'AMSTERDAM': 'NLAMS', 'FLUSHING': 'NLFLU', 'HEERENVEEN': 'NLHRV', 'IJMUIDEN': 'NLIJM', 'MOERDIJK': 'NLMOE', 'ROTTERDAM': 'NLRTM', 'VLISSINGEN': 'NLVLI', 'ALESUND': 'NOAES', 'ALTA': 'NOALF', 'ALSTAHAUG': 'NOALS', 'AUSTEVOLL': 'NOASV', 'BERGEN': 'NOBGO', 'BREVIK': 'NOBVK', 'EIDFJORD': 'NOEDF', 'EGERSUND': 'NOEGE', 'FLAM': 'NOFLA', 'FREDRIKSTAD': 'NOFRK', 'FLORO': 'NOFRO', 'GJEMNES': 'NOGJM', 'GEIRANGER': 'NOGNR', 'HALDEN': 'NOHAL', 'HAUGESUND': 'NOHAU', 'HOGSET': 'NOHOG', 'HONNINGSVAG': 'NOHVG', 'IKORNNES': 'NOIKR', 'KARMOY': 'NOKMY', 'KRISTIANSAND': 'NOKRS', 'KRISTIANSUND': 'NOKSU', 'KVINESDAL': 'NOKVD', 'LARVIK': 'NOLAR', 'LEIKANGER': 'NOLEK', 'LEKNES': 'NOLKN', 'LONGYEARBYEN': 'NOLYR', 'MALOY': 'NOMAY', 'MOLDE': 'NOMOL', 'MOSS': 'NOMSS', 'NARVIK': 'NONVK', 'OLDEN': 'NOOLD', 'ORKANGER': 'NOORK', 'OSLO': 'NOOSL', 'SALTEN': 'NOSAT', 'SAUDA': 'NOSAU', 'SUNNDALSORA': 'NOSUN', 'SVELGEN': 'NOSVE', 'STAVANGER': 'NOSVG', 'TANANGER': 'NOTAE', 'TROMSO': 'NOTOS', 'TRONDHEIM': 'NOTRD', 'AUCKLAND': 'NZAKL', 'BLUFF': 'NZBLU', 'CHRISTCHURCH': 'NZCHC', 'DUNEDIN': 'NZDUD', 'LYTTELTON': 'NZLYT', 'MARSDEN POINT': 'NZMAP', 'NAPIER': 'NZNPE', 'NELSON': 'NZNSN', 'PORT CHALMERS': 'NZPOE', 'TIMARU': 'NZTIU', 'TAURANGA': 'NZTRG', 'WELLINGTON': 'NZWLG', 'DUQM': 'OMDQM', 'KHASAB': 'OMKHS', 'MUSCAT': 'OMMCT', 'SALALAH': 'OMSLL', 'SOHAR': 'OMSOH', 'BALBOA': 'PABLB', 'CRISTOBAL': 'PACTB', 'COLON': 'PAONX', 'ALMIRANTE': 'PAPAM', 'PANAMA': 'PAPTY', 'RODMAN': 'PAROD', 'CALLAO': 'PECLL', 'PAITA': 'PEPAI', 'PISCO': 'PEPIO', 'SALAVERRY': 'PESVY', 'BORA-BORA': 'PFBOB', 'MOOREA': 'PFMOZ', 'PAPEETE': 'PFPPT', 'ALOTAU': 'PGGUR', 'LAE': 'PGLAE', 'LOSUIA': 'PGLSA', 'PORT MORESBY': 'PGPOM', 'RABAUL': 'PGRAB', 'BATANGAS, LUZON': 'PHBTG', 'CEBU': 'PHCEB', 'CAGAYAN DE ORO, MINDANAO': 'PHCGY', 'DAVAO, MINDANAO': 'PHDVO', 'GENERAL SANTOS': 'PHGES', 'MANILA NORTH HARBOUR': 'PHMNN', 'MANILA SOUTH HARBOUR': 'PHMNS', 'SUBIC': 'PHSFS', 'KARACHI-MUHAMMAD BIN QASIM': 'PKBQM', 'KARACHI': 'PKKHI', 'GDANSK': 'PLGDN', 'GDYNIA': 'PLGDY', 'SWINOUJSCIE': 'PLSWI', 'SZCZECIN': 'PLSZZ', 'PITCAIRN IS': 'PNPCN', 'SAN JUAN': 'PRSJU', 'AVEIRO': 'PTAVE', 'FIGUEIRA DA FOZ': 'PTFDF', 'FUNCHAL, MADEIRA': 'PTFNC', 'LEIXOES': 'PTLEI', 'LISBOA': 'PTLIS', 'PONTA DELGADA': 'PTPDL', 'PORTIMAO': 'PTPRM', 'SETUBAL': 'PTSET', 'SINES': 'PTSIE', 'SESIMBRA': 'PTSSB', 'TERCEIRA ISLAND': 'PTTER', 'KOROR': 'PWROR', 'CAACUPEMI ASUNCION': 'PYBCM', 'ENCARNACION PUERTO SAN JUAN': 'PYENO', 'CAACUPEMI PILAR': 'PYPIL', 'TERPORT VILLETA': 'PYTVT', 'PUERTO SEGURO FLUVIAL (VILLETA)': 'PYVLL', 'DOHA': 'QADOH', 'HAMAD': 'QAHMD', 'MESAIEED': 'QAMES', 'RAS LAFFAN': 'QARLF', 'POINTE DES GALETS': 'REPDG', 'POSSESSION': 'REPOS', 'AGIGEA': 'ROAGI', 'CONSTANTA': 'ROCND', 'GALATI': 'ROGAL', 'MANGALIA': 'ROMAG', 'ARKHANGELSK': 'RUARH', 'BALTIYSK': 'RUBLT', 'KRONSHTADT': 'RUKDT', 'KAVKAZ': 'RUKZP', 'SAINT PETERSBURG': 'RULED', 'NAKHODKA': 'RUNJK', 'NOVOROSSIYSK': 'RUNVS', 'PETROPAVLOVSK-KAMCHATSKIY': 'RUPKC', 'SLAVYANKA': 'RUSKA', 'SOCHI': 'RUSOC', "UST'-LUGA": 'RUULU', 'VLADIVOSTOK': 'RUVVO', 'ZARUBINO': 'RUZAR', 'AD DAMMAM': 'SADMM', 'JEDDAH': 'SAJED', 'JUBAIL': 'SAJUB', 'KING ABDULLAH PORT': 'SAKAC', 'NEOM': 'SANEO', 'YANBU AL-BAHR': 'SAYNB', 'HONIARA, GUADALCANAL IS': 'SBHIR', 'PORT VICTORIA': 'SCPOV', 'VICTORIA': 'SCVIC', 'PORT SUDAN': 'SDPZU', 'AHUS': 'SEAHU', 'GOTEBORG': 'SEGOT', 'GAVLE': 'SEGVX', 'HALMSTAD': 'SEHAD', 'HELSINGBORG': 'SEHEL', 'KARLSHAMN': 'SEKAN', 'NORRKOPING': 'SENRK', 'NYNASHAMN': 'SENYN', 'PITEA': 'SEPIT', 'SKELLEFTEA': 'SESFT', 'SKARHAMN': 'SESKM', 'SODERTALJE': 'SESOE', 'STOCKHOLM': 'SESTO', 'VISBY': 'SEVBY', 'SINGAPORE': 'SGSIN', 'JAMESTOWN': 'SHSHN', 'KOPER': 'SIKOP', 'FREETOWN': 'SLFNA', 'DAKAR': 'SNDKR', 'ZIGUINCHOR': 'SNZIG', 'BERBERA': 'SOBBO', 'KISMAYU': 'SOKMU', 'MOGADISHU': 'SOMGQ', 'PARAMARIBO': 'SRPBM', 'ACAJUTLA': 'SVAQJ', 'PHILIPSBURG': 'SXPHI', 'LATTAKIA': 'SYLTK', 'TARTUS': 'SYTTS', 'GRAND TURK ISLAND': 'TCGDT', 'PROVIDENCIALES': 'TCPLS', 'LOME': 'TGLFW', 'BANGKOK': 'THBKK', 'BANGKOK MODERN TERMINALS/BANGKOK': 'THBMT', 'PHUKET': 'THHKT', 'LAEM CHABANG': 'THLCH', 'LAT KRABANG': 'THLKR', 'PAT BANGKOK': 'THPAT', 'SIAM BANGKOK PORT': 'THSBP', 'SONGKHLA': 'THSGZ', 'THAI CONNECTIVITY TERMINAL': 'THTPT', 'KOH SAMUI': 'THUSM', 'DILI': 'TLDIL', 'LA GOULETTE NORD (HALQUELOUED)': 'TNLGN', 'RADES/TUNIS': 'TNRDS', 'SFAX': 'TNSFA', 'SOUSSE': 'TNSUS', 'TUNIS': 'TNTUN', "NUKU'ALOFA": 'TOTBU', 'ALANYA': 'TRALA', 'ALIAGA': 'TRALI', 'AVCILAR': 'TRAVC', 'ANTALYA': 'TRAYT', 'BANDIRMA': 'TRBDM', 'BESIKTAS': 'TRBTS', 'BODRUM': 'TRBXN', 'BOZCAADA': 'TRBZC', 'CANAKKALE': 'TRCKZ', 'DERINCE': 'TRDRC', 'EVYAP PORT': 'TREYP', 'GEBZE': 'TRGEB', 'GEMLIK': 'TRGEM', 'GIRESUN': 'TRGIR', 'ISKENDERUN': 'TRISK', 'ISTANBUL': 'TRIST', 'ISTINYE/BOSPHORUS': 'TRITY', 'IZMIR': 'TRIZM', 'LIMAS': 'TRLMA', 'MERSIN': 'TRMER', 'MARMARIS': 'TRMRM', 'SAMSUN': 'TRSSX', 'TEKIRDAG (ASYAPORT)': 'TRTEK', 'TUZLA': 'TRTUZ', 'TRABZON': 'TRTZX', 'YALOVA': 'TRYAL', 'YARIMCA': 'TRYAR', 'ZONGULDAK': 'TRZON', 'PORT-OF-SPAIN': 'TTPOS', 'POINT LISAS': 'TTPTS', 'SCARBOROUGH/TOBAGO': 'TTSCA', 'FUNAFUTI': 'TVFUN', 'KEELUNG': 'TWKEL', 'KAOHSIUNG': 'TWKHH', 'TAIPEI': 'TWTPE', 'TAICHUNG': 'TWTXG', 'DAR ES SALAAM': 'TZDAR', 'MTWARA': 'TZMYW', 'TANGA': 'TZTGT', 'ZANZIBAR': 'TZZNZ', 'CHORNOMORSK': 'UAILK', 'IZMAIL': 'UAIZM', 'ODESA': 'UAODS', 'RENI': 'UARNI', 'YALTA': 'UAYAL', 'YUZHNYY': 'UAYUZ', 'BALTIMORE': 'USBAL', 'BAR HARBOR': 'USBHB', 'BOSTON': 'USBOS', 'BROWNSVILLE': 'USBRO', 'CHARLESTON': 'USCHS', 'PT ANGELES': 'USCLM', 'CAPE CANAVERAL': 'USCPV', 'DUTCH HARBOR': 'USDUT', 'DAVISVILLE': 'USDVV', 'ELLISVILLE': 'USEVE', 'NEWARK': 'USEWR', 'KEY WEST': 'USEYW', 'FORT LAUDERDALE': 'USFLL', 'GULFPORT': 'USGPT', 'HONOLULU': 'USHNL', 'HOUSTON': 'USHOU', 'WILMINGTON, DE': 'USILG', 'WILMINGTON, NC': 'USILM', 'HILO': 'USITO', 'JACKSONVILLE': 'USJAX', 'KAWAIHAE': 'USKWH', 'LOS ANGELES': 'USLAX', 'LONG BEACH': 'USLGB', 'MIAMI': 'USMIA', 'MOBILE': 'USMOB', 'NEW ORLEANS': 'USMSY', 'NAWILIWILI': 'USNIJ', 'NEWPORT': 'USNPO', 'PORT HUENEME': 'USNTD', 'NEW YORK': 'USNYC', 'OAKLAND': 'USOAK', 'NORFOLK': 'USORF', 'PORTLAND, OR': 'USPDX', 'PORT EVERGLADES': 'USPEF', 'PHILADELPHIA': 'USPHL', 'PORTLAND, ME': 'USPWM', 'SAVANNAH': 'USSAV', 'SEATTLE': 'USSEA', 'SAN FRANCISCO': 'USSFO', 'TACOMA': 'USTIW', 'TAMPA': 'USTPA', 'UNALASKA': 'USUAA', 'MONTEVIDEO': 'UYMVD', 'NUEVA PALMIRA': 'UYNVP', 'PUNTA DEL ESTE': 'UYPDP', 'CAMPDEN PARK': 'VCCRP', 'KINGSTOWN, ST VINCENT': 'VCKTN', 'EL TABLAZO/MARACAIBO L': 'VEETV', 'GUARANAO BAY': 'VEGUB', 'LA GUAIRA': 'VELAG', 'PUERTO CABELLO': 'VEPBL', 'PUERTO LA CRUZ': 'VEPCZ', 'N. SOUND/VIRGIN GORDA': 'VGNSX', 'ROAD TOWN, TORTOLA': 'VGRAD', 'CHARLOTTE AMALIE, ST THOMAS': 'VICHA', 'SAINT THOMAS': 'VISTT', 'DA-NANG': 'VNDAD', 'HAIPHONG': 'VNHPH', 'NHA TRANG': 'VNNHA', 'PHUOC LONG': 'VNPHG', 'HO CHI MINH CITY': 'VNSGN', 'QUINHON': 'VNUIH', 'VUNG TAU': 'VNVUT', 'PORT VILA': 'VUVLI', 'APIA': 'WSAPW', 'ADEN': 'YEADE', 'HODEIDAH': 'YEHOD', 'MUKALLA': 'YEMKX', 'LONGONI': 'YTLON', 'CAPE TOWN': 'ZACPT', 'DURBAN': 'ZADUR', 'EAST LONDON': 'ZAELS', 'MOSSEL BAY': 'ZAMZY', 'PORT ELIZABETH': 'ZAPLZ', 'RICHARDS BAY': 'ZARCB', 'COEGA': 'ZAZBA'}

    if isinstance(port_name, str):
        return port_codes.get(port_name.upper(), 'UNKNOWN')
    return 'UNKNOWN'



# 드래그 앤 드롭 핸들러
def drop_tsteam(event):
    global tsteam_file
    tsteam_file = root.tk.splitlist(event.data)[0]
    tsteam_label.config(text=f"TSTEAM Data: {tsteam_file}")

def drop_itps(event):
    global itps_file
    itps_file = root.tk.splitlist(event.data)[0]
    itps_label.config(text=f"ITPS Data: {itps_file}")

def process_files():
    pnit_vessel = pnit_vessel_entry.get()
    pnit_pod = pnit_pod_entry.get()
    pnc_vessel = pnc_vessel_entry.get()
    pnc_pod = pnc_pod_entry.get()

    # 파일 로드
    tsteam_data = pd.read_excel(tsteam_file, sheet_name=None)
    itps_data = pd.read_excel(itps_file)

    # TSTEAM_DATA 처리
    ts_list = []
    for sheet_name, df in tsteam_data.items():
        if 'Cntr No.' in df.columns or 'Container Number' in df.columns:
            cntr_col = 'Cntr No.' if 'Cntr No.' in df.columns else 'Container Number'
            for _, row in df.iterrows():
                por_value = row['POR'] if 'POR' in df.columns else row['POL']
                fpod_fnd_value = row['FPOD'] if 'FPOD' in df.columns else (row['FND'] if 'FND' in df.columns else '')
                ts_list.append({
                    'Cntr No.': row[cntr_col],
                    'POR': por_value,
                    'POL': row['POL'],
                    'POD': row['POD'],
                    'FPOD': fpod_fnd_value,
                    'REMARK': row['REMARK'] if 'REMARK' in df.columns else ''
                })

    # TS_LIST DataFrame 생성
    ts_df = pd.DataFrame(ts_list)

    # ITPS_DATA 처리 및 대체
    if 'Equipment Number' in itps_data.columns:
        def replace_values(row):
            match = ts_df[ts_df['Cntr No.'] == row['Equipment Number']]
            if not match.empty:
                row['Origin Load Port'] = match.iloc[0]['POR']
                row['Next POD'] = match.iloc[0]['POD']
                # FPOD 정보를 사용하지 않고 Discharge Port 값을 convert_to_port_code 함수로 변환
                discharge_port = row['Discharge Port']
                if pd.notnull(discharge_port):
                    port_name_parts = discharge_port.split(',')
                    port_name = ','.join(port_name_parts[:-1]).strip()
                else:
                    port_name = ''
                new_discharge_port = convert_to_port_code(port_name)
                if new_discharge_port != 'UNKNOWN':
                    row['Discharge Port'] = new_discharge_port
                row['Active'] = match.iloc[0]['REMARK']

            # if 'SOC' in itps_data.columns:
            #     soc_value = str(row.get('SOC', ''))
            #     if soc_value == 'Y' and 'SOC' not in row['Remark']:
            #         row['Remark'] = (row['Remark'] + ', SOC').strip()

            return row

        itps_data = itps_data.apply(replace_values, axis=1)


    # VGM Weight가 비어있다면 Weight 값을 복사하여 반올림 후 소수점 부분 없애기
    itps_data['VGM Weight'] = itps_data['VGM Weight'].fillna(itps_data['Weight']).fillna(0).apply(lambda x: round(x))

    itps_data['VGM Unit'].fillna('KGS', inplace=True)

    # VGM Unit이 LBS라면 KGS로 변환
    def convert_lbs_to_kgs(row):
        if row['VGM Unit'] == 'LBS':
            row['VGM Weight'] = round(row['VGM Weight'] * 0.453592)
            row['VGM Unit'] = 'KGS (from LBS)'
        return row

    itps_data = itps_data.apply(convert_lbs_to_kgs, axis=1)

    # Discharge Port의 값이 'KR'로 시작하지 않는데 빈칸일 경우 처리
    if 'Discharge Port' in itps_data.columns:
        def fill_discharge_port(row):
            if not pd.isnull(row['Discharge Port']) and not row['Discharge Port'].startswith('KR'):
                if pnit_vessel:
                    row['Loading Terminal'] = 'PUSAN NEWPORT INTERNATIONAL TERMINAL-PNIT'
                    row['Loaded Vessel/Voyage'] = pnit_vessel
                    row['Next POD'] = pnit_pod
                    row['Discharge Port'] = pnit_pod
                elif pnc_vessel:
                    row['Loading Terminal'] = 'PUSAN NEWPORT COMPANY LIMITED'
                    row['Loaded Vessel/Voyage'] = pnc_vessel
                    row['Next POD'] = pnc_pod
                    row['Discharge Port'] = pnc_pod
            return row

        itps_data = itps_data.apply(fill_discharge_port, axis=1)

    # SUMMARY 데이터 생성

    if 'Discharge Port' in itps_data.columns:
        def classify_pod(row):
            if row['Full/Empty'] != 'E':  # Full/Empty가 'E'일 경우에만 아래 코드 실행
                if row['Discharge Port'] in ['KRPUS', 'KRINC', 'KRKAN']:
                    return row['Discharge Port']
                else:
                    return 'TS'
            return None  # 'E'가 아닌 경우 None 반환

        itps_data['L/T'] = itps_data.apply(classify_pod, axis=1)
        itps_data['L/T'].fillna(itps_data['L/T'], inplace=True)  # None을 원래 L/T 값으로 대체
    else:
        itps_data['L/T'] = 'TS'

    if 'Type/Size' in itps_data.columns:
        summary_data = itps_data.pivot_table(index='Type/Size', columns='L/T', aggfunc='size', fill_value=0).reset_index()
        summary_data['Empty'] = itps_data[itps_data['Full/Empty'] == 'E'].groupby('Type/Size').size().reindex(summary_data['Type/Size'], fill_value=0).values
        summary_data['Total'] = summary_data.sum(numeric_only=True, axis=1).astype(int)
        summary_data.loc['Total'] = summary_data.sum(numeric_only=True, axis=0).astype(int)
        summary_data.loc['Total', 'Type/Size'] = 'Total'

        # Ensure all numeric values are integers
        for col in summary_data.columns:
            if summary_data[col].dtype != 'object':
                summary_data[col] = summary_data[col].astype(int)

        # Reorder columns to ensure 'TS', 'Empty', 'Total' are in the correct order
        cols = summary_data.columns.tolist()
        if 'TS' in cols and 'Empty' in cols and 'Total' in cols:
            cols.remove('Empty')
            ts_index = cols.index('TS')
            cols.insert(ts_index + 1, 'Empty')
            summary_data = summary_data[cols]
    else:
        summary_data = pd.DataFrame({'Type/Size': ['No Type/Size Data'], 'TS': [0], 'Empty': [0], 'Total': [0]})


    # 모든 숫자 열을 정수로 변환
    for col in summary_data.columns:
        if summary_data[col].dtype != 'object':
            summary_data[col] = summary_data[col].astype(int)

    itps_data.rename(columns={'Next POD': 'Stow Code', 'Active': 'Remark'}, inplace=True)

    if 'Remark' not in itps_data.columns:
        itps_data['Remark'] = ''

    # Convert 'IMO Class' column to string explicitly, and handle NaN values by converting them to empty strings
    itps_data['IMO Class'] = itps_data['IMO Class'].astype(str).replace('nan', '')

    # Define the function to add '직반출 위험물' to the Remark column
    def add_risk_remark(row):
        imo_class = row['IMO Class'].strip()  # Ensure IMO Class is treated as string and strip any spaces
        reefer_temp = row['Reefer Temp.']
        # Check if IMO Class is exactly '2', '2.1', or '2.2' or both IMO Class and Reefer Temp. have values
        if imo_class in ['2', '2.1', '2.2'] or (imo_class and pd.notna(reefer_temp)):
            row['Remark'] = row['Remark'] + " " +'직반출 위험물'
        return row
    
    def add_soc_remark(row):
        soc_itps = row['SOC']

        imo_class = row['IMO Class'].strip()  # Ensure IMO Class is treated as string and strip any spaces
        reefer_temp = row['Reefer Temp.']
        # Check if IMO Class is exactly '2', '2.1', or '2.2' or both IMO Class and Reefer Temp. have values
        # if imo_class in ['2', '2.1', '2.2'] or (imo_class and pd.notna(reefer_temp)):
        #     row['Remark'] = '직반출 위험물'

        if row['SOC'] == 'Y':
            row['Remark'] = row['Remark'] + " "  + 'SOC'

        elif row['SOC'] == 'Y' and (imo_class in ['2', '2.1', '2.2'] or (imo_class and pd.notna(reefer_temp))):
            row['Remark'] = row['Remark']+ " " + 'SOC, 직반출 위험물'
    
        return row

    # Ensure 'Remark' column is in the dataframe and set to an empty string if it doesn't exist
    if 'Remark' not in itps_data.columns:
        itps_data['Remark'] = ''

    # Convert 'Remark' column to string to avoid AttributeError
    itps_data['Remark'] = itps_data['Remark'].astype(str).replace('nan', '')

    # Apply the function to the dataframe
    itps_data = itps_data.apply(add_risk_remark, axis=1)
    itps_data = itps_data.apply(add_soc_remark, axis=1)

    # `BL No.<Booking No.>` 열의 값에서 '<'를 포함한 뒷부분 값을 모두 버리도록 처리
    if 'BL No.<Booking No.>' in itps_data.columns:
        itps_data['BL No.<Booking No.>'] = itps_data['BL No.<Booking No.>'].apply(lambda x: x.split('<')[0] if pd.notnull(x) else x)

    # Discharging Terminal, Loading Terminal 값 변경
    if 'Discharging Terminal' in itps_data.columns:
        itps_data['Discharging Terminal'] = itps_data['Discharging Terminal'].replace({
            'PUSAN NEWPORT INTERNATIONAL TERMINAL-PNIT': 'PNIT',
            'PUSAN NEWPORT COMPANY LIMITED': 'PNC'
        })

    if 'Loading Terminal' in itps_data.columns:
        itps_data['Loading Terminal'] = itps_data['Loading Terminal'].replace({
            'PUSAN NEWPORT INTERNATIONAL TERMINAL-PNIT': 'PNIT',
            'PUSAN NEWPORT COMPANY LIMITED': 'PNC'
        })

    # FEDDERS 단어가 들어 있는 경우 Loading Terminal과 Stow Code 값을 설정
        def update_terminal_and_stow_code(row):
            if 'FEEDERS' in str(row['Loaded Vessel/Voyage']):
                if row['Loading Terminal'] == 'PUSAN NEWPORT COMPANY LIMITED':
                    row['Loading Terminal'] = 'PNC'
                elif row['Loading Terminal'] == 'PUSAN NEWPORT INTERNATIONAL TERMINAL-PNIT':
                    row['Loading Terminal'] = 'PNIT'
                else:
                    row['Loading Terminal'] = row['Loading Terminal']

                row['Stow Code'] = row['Discharge Port']

            return row

    # Apply the function to the dataframe
    itps_data = itps_data.apply(update_terminal_and_stow_code, axis=1)

    # 특정 조건을 만족하는 경우 Loaded Vessel/Voyage 값을 채우는 함수
    def fill_loaded_vessel_voyage(row):
        if pd.isna(row['Loaded Vessel/Voyage']):
            matching_rows = [item for item in port_svc.items() if item[0][0] == row['Loading Terminal'] and item[0][1] == row['Stow Code']]
            if matching_rows:
                service_name = matching_rows[0][1]
                msc_schedule = [(tree.set(child, 'Arrival at Berth'), tree.set(child, 'Vessel'), tree.set(child, 'Vessel Voyage')) 
                                for child in tree.get_children() 
                                if tree.set(child, 'Service') == service_name]
                msc_schedule = sorted([(datetime.strptime(date, "%Y-%m-%d"), vessel, voyage) 
                                       for date, vessel, voyage in msc_schedule 
                                       if datetime.strptime(date, "%Y-%m-%d") >= datetime.now()])
                if len(msc_schedule) > 1:
                    row['Loaded Vessel/Voyage'] = f"{msc_schedule[1][1]}/{msc_schedule[1][2]}"
        return row

    # Apply the function to fill Loaded Vessel/Voyage
    itps_data = itps_data.apply(fill_loaded_vessel_voyage, axis=1)

    # 생성할 열 이름 및 순서
    columns_order = [
        'Equipment Number', 'BL No.<Booking No.>', 'Discharge Vessel/Voyage', 'Discharging Terminal', 
        'Loaded Vessel/Voyage', 'Loading Terminal', 'Origin Load Port', 'Stow Code', 'Discharge Port', 
        'Type/Size', 'Full/Empty', 'Weight', 'VGM Weight', 'VGM Unit', 'Reefer or IMO', 'Reefer Temp.', 
        'Dangerous Cargo', 'IMO Class', 'UN Number', 'Vet. Control', 'Remark'
    ]

    # 필요한 열만 남기고 순서대로 정렬
    itps_data = itps_data.reindex(columns=columns_order, fill_value='')

    # SOC 탭 데이터 생성
    soc_data = itps_data[itps_data['Remark'].str.contains('SOC', na=False)]
    soc_data_selected = soc_data[['BL No.<Booking No.>','Equipment Number', 'Origin Load Port', 'Discharge Port', 'Full/Empty', 'Type/Size']]
    soc_data_total = pd.DataFrame([['Total', '','', '', '', soc_data_selected.shape[0]]], columns=soc_data_selected.columns)

    # 직반출 위험물 탭 데이터 생성
    # dangerous_goods_data = itps_data[itps_data['Remark'].str.contains('직반출 위험물', na=False)]
    # dangerous_goods_data_selected = dangerous_goods_data[['Equipment Number', 'Origin Load Port', 'Discharge Port', 'IMO Class', 'UN Number', 'Reefer Temp.']]

    dangerous_goods_data = itps_data[itps_data['Remark'].str.contains('직반출 위험물', na=False)]
    dangerous_goods_data_selected = dangerous_goods_data[['BL No.<Booking No.>', 'Equipment Number', 'Full/Empty', 'Origin Load Port', 'Discharge Port', 'IMO Class', 'UN Number', 'Reefer Temp.']]

    # NaN 값을 빈칸으로 변경
    dangerous_goods_data_selected['Reefer Temp.'] = dangerous_goods_data_selected['Reefer Temp.'].fillna('')

    dangerous_goods_data_total = pd.DataFrame([['Total', '', '', '', '','','', dangerous_goods_data_selected.shape[0]]], columns=dangerous_goods_data_selected.columns)

    # 파일 이름 생성
    tsteam_filename = os.path.splitext(os.path.basename(tsteam_file))[0]
    output_path = os.path.join(os.path.dirname(tsteam_file), f"{tsteam_filename} merged file.xlsx")

    # 파일 저장
    with pd.ExcelWriter(output_path) as writer:
        itps_data.to_excel(writer, index=False, sheet_name='ITPS_DATA')
        summary_data.to_excel(writer, index=False, sheet_name='SUMMARY')
        soc_data_selected.to_excel(writer, index=False, sheet_name='SOC')
        soc_data_total.to_excel(writer, index=False, header=False, startrow=soc_data_selected.shape[0] + 1, sheet_name='SOC')
        dangerous_goods_data_selected.to_excel(writer, index=False, sheet_name='직반출 위험물')
        dangerous_goods_data_total.to_excel(writer, index=False, header=False, startrow=dangerous_goods_data_selected.shape[0] + 1, sheet_name='직반출 위험물')

    # 엑셀 파일 열기
    wb = openpyxl.load_workbook(output_path)
    ws_itps = wb['ITPS_DATA']

    # A2 셀에 값이 없을 경우 2행 전체 삭제
    if ws_itps['A2'].value is None:
        ws_itps.delete_rows(2)

    # 스타일 정의
    pnc_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # Light Blue
    pnit_fill = PatternFill(start_color="FFFFE0", end_color="FFFFE0", fill_type="solid")  # Light Yellow
    feeder_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")  # Light Purple

    # 스타일 적용 함수
    def apply_style(sheet, column_letter):
        for row in range(2, sheet.max_row + 1):
            cell = sheet[f"{column_letter}{row}"]
            if cell.value == 'PNC':
                cell.fill = pnc_fill
            elif cell.value == 'PNIT':
                cell.fill = pnit_fill
            elif cell.value == '타부두':
                cell.fill = feeder_fill

    # Discharging Terminal에 스타일 적용
    discharging_terminal_col = None
    loading_terminal_col = None

    for cell in ws_itps[1]:
        if cell.value == 'Discharging Terminal':
            discharging_terminal_col = cell.column_letter
        elif cell.value == 'Loading Terminal':
            loading_terminal_col = cell.column_letter

    if discharging_terminal_col:
        apply_style(ws_itps, discharging_terminal_col)
    if loading_terminal_col:
        apply_style(ws_itps, loading_terminal_col)

    # 모든 시트에 대해 열 너비 자동 조정
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

    # 저장
    wb.save(output_path)

    # 통계 정보 출력
    # container_count = len(ts_df['Cntr No.'].unique())
    container_count = len(ts_df[ts_df['Cntr No.'].str.len() == 11]['Cntr No.'].unique())
    full_count = itps_data[itps_data['Full/Empty'] == 'F'].shape[0]
    empty_count = itps_data[itps_data['Full/Empty'] == 'E'].shape[0]

    result_label.config(text=f"Container count in TSTEAM_DATA: {container_count}\n"
                             f"F count in Full/Empty in ITPS_DATA: {full_count}\n"
                             f"E count in Full/Empty in ITPS_DATA: {empty_count}\n"
                             f"Merged file saved as: {output_path}")

    # Summary 데이터를 Summary 탭에서 표시
    show_summary(summary_data)

    # SOC 데이터와 직반출 위험물 데이터를 각각의 새로운 탭에 표시
    show_data_in_tab('SOC Data', soc_data_selected)
    show_data_in_tab('직반출 위험물 Data', dangerous_goods_data_selected)

def show_data_in_tab(tab_name, data):
    new_tab = ttk.Frame(tab_control)
    tab_control.add(new_tab, text=tab_name)
    tab_control.pack(expand=1, fill='both')

    for widget in new_tab.winfo_children():
        widget.destroy()

    # Treeview 설정
    columns = list(data.columns)
    tree = ttk.Treeview(new_tab, columns=columns, show='headings')

    # 각 열의 제목 설정
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor='center', width=100)

    # 데이터 삽입
    for index, row in data.iterrows():
        tree.insert('', 'end', values=list(row))

    tree.pack(expand=True, fill='both')

    copy_button = tk.Button(new_tab, text="Copy to Clipboard", command=lambda: copy_treeview_to_clipboard(tree, columns))
    copy_button.pack(pady=10)

def show_summary(summary_data):
    for widget in summary_tab.winfo_children():
        widget.destroy()

    # Treeview 설정
    columns = list(summary_data.columns)
    tree = ttk.Treeview(summary_tab, columns=columns, show='headings')

    # 각 열의 제목 설정
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor='center', width=100)

    # 데이터 삽입
    for index, row in summary_data.iterrows():
        tree.insert('', 'end', values=list(row))

    tree.pack(expand=True, fill='both')

    copy_button = tk.Button(summary_tab, text="Copy to Clipboard", command=lambda: copy_treeview_to_clipboard(tree, columns))
    copy_button.pack(pady=10)

def copy_treeview_to_clipboard(tree, columns):
    tree_data = []
    tree_data.append('\t'.join(columns))  # Adding headers
    for item in tree.get_children():
        row = tree.item(item)['values']
        row_str = '\t'.join(map(str, row))
        tree_data.append(row_str)

    clipboard_data = '\n'.join(tree_data)
    root.clipboard_clear()
    root.clipboard_append(clipboard_data)
    root.update()

def copy_to_clipboard(data):
    root.clipboard_clear()
    root.clipboard_append(data)
    root.update()

pastel_colors = [
    "#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF",
    "#CBA6E1", "#FFD1DC", "#F0C27B", "#FFB347", "#A0E7E5"
]

def color_treeview(tree):
    children = tree.get_children('')
    if not children:
        return

    service_colors = {}
    current_color = 0

    for child in children:
        service = tree.set(child, 'Service')
        if service not in service_colors:
            service_colors[service] = pastel_colors[current_color % len(pastel_colors)]
            current_color += 1
        tree.tag_configure(service, background=service_colors[service])
        tree.item(child, tags=(service,))

def paste_from_clipboard():
    # 클립보드에서 텍스트 가져오기
    clipboard_text = pyperclip.paste()

    # 헤더 정의
    headers = ['Vessel', 'Vessel Voyage', 'Arrival at Berth', 'Departure', 'Service']

    # Treeview 초기화
    for child in tree.get_children():
        tree.delete(child)

    # 클립보드에서 가져온 데이터를 행별로 처리
    for line in clipboard_text.strip().split('\n'):
        # 데이터 분할
        data = line.split('\t')
        # Treeview에 데이터 추가
        tree.insert("", "end", values=(data[0], data[2], data[3], data[4], data[5]))

    # 같은 Service를 같은 색으로 색칠
    color_treeview(tree)

def export_treeview_to_excel(tree, filename):
    # 엑셀 워크북 생성
    wb = Workbook()
    ws = wb.active

    # Treeview의 컬럼 헤더 가져오기
    columns = tree["columns"]
    ws.append(columns)

    # Treeview의 데이터 가져오기
    for child in tree.get_children():
        row_values = tree.item(child)["values"]
        row_index = len(ws['A']) + 1  # 현재 행 번호

        for col_index, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_index, column=col_index, value=value)
            # Treeview의 배경색을 가져와 엑셀 셀에 적용
            tags = tree.item(child)["tags"]
            if tags:
                bg_color = tree.tag_configure(tags[0])["background"]
                if bg_color:
                    argb_color = hex_to_argb(bg_color)
                    cell.fill = PatternFill(start_color=argb_color, end_color=argb_color, fill_type="solid")

    # 엑셀 파일 저장
    wb.save(filename)
    print(f"File saved: {filename}")




def save_treeview_to_excel():
    # 파일 저장 대화 상자
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="live_schedule.xlsx")
    if file_path:
        export_treeview_to_excel(tree, file_path)

def hex_to_argb(hex_color):
    if hex_color.startswith("#"):
        hex_color = hex_color[1:]
    if len(hex_color) == 6:
        hex_color = "FF" + hex_color
    return hex_color

# Tkinter 설정
root = TkinterDnD.Tk()
root.title("인바운드 컨버터")
root.geometry('800x600')

# 탭 컨트롤 생성
tab_control = ttk.Notebook(root)
root_tab = ttk.Frame(tab_control)
summary_tab = ttk.Frame(tab_control)

tab_control.add(root_tab, text='Root')
tab_control.add(summary_tab, text='Summary')
tab_control.pack(expand=1, fill='both')

# 중앙 정렬을 위한 레이아웃 프레임
center_frame = tk.Frame(root_tab)
center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

# PNIT 관련 입력 필드
pnit_label = tk.Label(center_frame, text="PNIT")
pnit_label.grid(row=0, column=0, padx=10, pady=5, sticky='e')

pnit_vessel_label = tk.Label(center_frame, text="VESSEL:")
pnit_vessel_label.grid(row=0, column=1, padx=10, pady=5, sticky='e')
pnit_vessel_entry = tk.Entry(center_frame)
pnit_vessel_entry.grid(row=0, column=2, padx=10, pady=5)

pnit_pod_label = tk.Label(center_frame, text="POD:")
pnit_pod_label.grid(row=0, column=3, padx=10, pady=5, sticky='e')
pnit_pod_entry = tk.Entry(center_frame)
pnit_pod_entry.grid(row=0, column=4, padx=10, pady=5)

# PNC 관련 입력 필드
pnc_label = tk.Label(center_frame, text="PNC")
pnc_label.grid(row=1, column=0, padx=10, pady=5, sticky='e')

pnc_vessel_label = tk.Label(center_frame, text="VESSEL:")
pnc_vessel_label.grid(row=1, column=1, padx=10, pady=5, sticky='e')
pnc_vessel_entry = tk.Entry(center_frame)
pnc_vessel_entry.grid(row=1, column=2, padx=10, pady=5)

pnc_pod_label = tk.Label(center_frame, text="POD:")
pnc_pod_label.grid(row=1, column=3, padx=10, pady=5, sticky='e')
pnc_pod_entry = tk.Entry(center_frame)
pnc_pod_entry.grid(row=1, column=4, padx=10, pady=5)

# TSTEAM 파일 드래그 앤 드롭 영역
tsteam_frame = tk.Frame(center_frame, width=350, height=100, bg='lightgrey')
tsteam_frame.grid(row=2, column=0, columnspan=5, pady=10)
tsteam_frame.drop_target_register(DND_FILES)
tsteam_frame.dnd_bind('<<Drop>>', drop_tsteam)

tsteam_label = tk.Label(tsteam_frame, text="Drop TSTEAM Data File Here")
tsteam_label.pack(padx=10, pady=10)

# ITPS 파일 드래그 앤 드롭 영역
itps_frame = tk.Frame(center_frame, width=350, height=100, bg='lightgrey')
itps_frame.grid(row=3, column=0, columnspan=5, pady=10)
itps_frame.drop_target_register(DND_FILES)
itps_frame.dnd_bind('<<Drop>>', drop_itps)

itps_label = tk.Label(itps_frame, text="Drop ITPS Data File Here")
itps_label.pack(padx=10, pady=10)

# 파일 처리 버튼
process_button = tk.Button(center_frame, text="Process Files", command=process_files, width=20, height=2)
process_button.grid(row=4, column=0, columnspan=5, pady=20)

# 결과 출력 라벨
result_label = tk.Label(root_tab, text="", anchor='w', justify='left')
result_label.place(x=50, y=450)

# Summary 탭 설정
summary_text_widget = tk.Text(summary_tab, wrap='none')
summary_text_widget.pack(expand=True, fill='both')

scrollbar_y = tk.Scrollbar(summary_text_widget, orient='vertical', command=summary_text_widget.yview)
scrollbar_y.pack(side='right', fill='y')
summary_text_widget.configure(yscrollcommand=scrollbar_y.set)

scrollbar_x = tk.Scrollbar(summary_text_widget, orient='horizontal', command=summary_text_widget.xview)
scrollbar_x.pack(side='bottom', fill='x')
summary_text_widget.configure(xscrollcommand=scrollbar_x.set)

#####================================ MSC LIVE SCHEDULE TAB

def sort_treeview(tree, col, reverse, sort_type):
    data = [(tree.set(child, col), child) for child in tree.get_children('')]
    
    if sort_type == 'date':
        data.sort(key=lambda x: datetime.strptime(x[0], "%Y-%m-%d"), reverse=reverse)
    else:
        data.sort(reverse=reverse)
    
    for index, (val, child) in enumerate(data):
        tree.move(child, '', index)
    
    tree.heading(col, command=lambda: sort_treeview(tree, col, not reverse, sort_type))

msc_ls = ttk.Frame(tab_control)
msc_ls.pack(fill="both", expand=True)
tab_control.add(msc_ls, text="MSC Live Schedule")

headers = ['Vessel', 'Vessel Voyage', 'Arrival at Berth', 'Departure', 'Service']

# Treeview 생성
tree = ttk.Treeview(msc_ls, columns=headers, show="headings")
for header in headers:
    if header == 'Arrival at Berth':
        tree.heading(header, text=header, command=lambda _col=header: sort_treeview(tree, _col, False, 'date'))
    else:
        tree.heading(header, text=header, command=lambda _col=header: sort_treeview(tree, _col, False, 'text'))

# 스크롤바 추가
scrollbar = ttk.Scrollbar(msc_ls, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

tree.pack(fill="both", expand=True)

tree.column(0, width=50)  # Vessel
tree.column(1, width=150)  # Vessel Voyage
tree.column(2, width=150)  # Arrival at Berth
tree.column(3, width=100)  # Departure
tree.column(4, width=100)  # Service

# 붙여넣기 버튼 생성
paste_button = ttk.Button(msc_ls, text="Paste from Clipboard", command=paste_from_clipboard)
paste_button.pack()

# MSC Live Schedule 탭에 엑셀로 저장 버튼 추가
save_button = tk.Button(msc_ls, text="Save to Excel", command=save_treeview_to_excel)
save_button.pack(pady=10)

# 새로운 Port_Svc 탭 추가
port_svc_tab = ttk.Frame(tab_control)
tab_control.add(port_svc_tab, text='Port_Svc')

# Treeview 생성
columns = ('PORT', 'POD', 'DIRECTION', 'SERVICE')
tree_port_svc = ttk.Treeview(port_svc_tab, columns=columns, show='headings')

# 각 컬럼의 헤더 설정
for col in columns: 
    tree_port_svc.heading(col, text=col)

# 데이터 추가
for (tml, port, direction), service in port_svc.items():
    tree_port_svc.insert("", "end", values=(tml, port, direction, service))

# 스크롤바 추가
scrollbar = ttk.Scrollbar(port_svc_tab, orient="vertical", command=tree_port_svc.yview)
tree_port_svc.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

tree_port_svc.pack(fill="both", expand=True)

root.mainloop()
