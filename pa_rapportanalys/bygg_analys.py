"""Resultatanalys PA (Electrolux Professional Sweden AB, bolag 00047).

Bygger en arbetsbok med levande formler mot JDE-uttaget R55EL121 (CSV) och
AARO-mappningsfilen (Sheet1 och JDE-flik).

Anrop: python3 -I bygg_analys.py <R55EL121.csv> <Aaro_mapping.xlsx> <ut.xlsx>
"""
import csv
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

CSV, MAPX, UT = sys.argv[1], sys.argv[2], sys.argv[3]

MANAD = 9          # rapportmånad
AR = 2026
SENA_VER = [26000344, 26000345]   # JE bokade efter AARO-uttaget (batchdatum 2026-10-06)

FET = Font(bold=True)
KURSIV = Font(italic=True)
TKR = '#,##0;-#,##0;0'
KR = '#,##0.00;-#,##0.00;0.00'
PROC = '0.0%'
DATUM = 'yyyy-mm-dd'

MANADSNAMN = ['jan', 'feb', 'mar', 'apr', 'maj', 'jun', 'jul', 'aug', 'sep',
              'okt', 'nov', 'dec']


# ---------------------------------------------------------------- inläsning
def las_csv(p):
    with open(p, encoding='utf-8-sig', newline='') as f:
        rader = list(csv.reader(f, delimiter=';'))
    hdr = [h.strip().replace('\xa0', '') for h in rader[4]]
    ix = {h: i for i, h in enumerate(hdr) if h}
    ut = []
    for n, r in enumerate(rader[5:], start=6):
        r = r + [''] * (len(hdr) - len(r))
        def g(k):
            return r[ix[k]].strip()
        def tal(s):
            s = s.replace(' ', '').replace(',', '.')
            return float(s) if s not in ('', '.') else 0.0
        ut.append({
            'rad': n,
            'mcu': g('Centro Costo').lstrip('_'),
            'obj': int(g('Konto')),
            'sub': g('Und- kto').lstrip('_'),
            'datum': datetime.strptime(g('HB- datum'), '%y-%m-%d'),
            'belopp': tal(g('Belopp')),
            'doktyp': g('Dok- typ'),
            'doknr': int(g('Dokument- nummer')),
            'batch': int(g('Batch- nummer')),
            'batchdatum': datetime.strptime(g('Batch- datum'), '%y-%m-%d'),
            'anv': g('Användar- id'),
            'beskr': g('Beskrivning söknamn'),
            'forkl': g('Förklaring/ kommentar'),
            'soknamn': g('Söknamn'),
            'valkod': g('Val- kod'),
            'valbelopp': tal(g('Valuta- belopp')),
        })
    return ut


data = las_csv(CSV)
for d in data:
    d['konto'] = f"{d['mcu']}.{d['obj']}.{d['sub']}"

src = openpyxl.load_workbook(MAPX, data_only=True)
s1 = src['Sheet1']
aaro_rows = []
for r in s1.iter_rows(min_row=2, values_only=True):
    if r[0] is None:
        continue
    aaro_rows.append([v.strip() if isinstance(v, str) else v for v in r])
jde_tab = []
for r in src['JDE 2609'].iter_rows(min_row=2, values_only=True):
    if r[0] is None:
        continue
    jde_tab.append([str(r[0]).strip(), (r[1] or '').strip(), r[2], str(r[3]).strip()])

beskr_aaro = {}
for r in aaro_rows:
    beskr_aaro.setdefault(r[1], r[2])

# ------------------------------------------------------------- mappning
mapp = {}   # konto -> (beskrivning, aaro, källa, current)
for k, b, cur, a in jde_tab:
    kalla = 'JDE 2609-fliken'
    if a == '#N/A':
        a, kalla = '(ej mappad)', 'JDE 2609-fliken visar #N/A'
    mapp[k] = [b, a, kalla, cur]
# AARO Sheet1 bär konton som saknas i JDE-fliken
for r in aaro_rows:
    k = f'{r[7]}.{r[9]}.{r[11]}'
    if k not in mapp:
        mapp[k] = [r[12], r[1], 'AARO Sheet1 (saknas i JDE 2609-fliken)', None]
# konton i uttaget utan mappning: komplettera efter kostnadsställe och konto
konto_beskr = {}
for d in data:
    konto_beskr.setdefault(d['konto'], d['beskr'])
regel = {('0473704', 2100): '4511', ('0473705', 2100): 'MARKET60',
         ('0470000', 8401): '1440'}
for k in sorted({d['konto'] for d in data} - set(mapp)):
    mcu, obj, sub = k.split('.')
    a = regel.get((mcu, int(obj)))
    if a:
        mapp[k] = ['', a, 'Komplettering efter kostnadsställe/konto', None]
    else:
        mapp[k] = ['', '(ej mappad)', 'Saknas i mappningsfilen', None]

# ------------------------------------------------------------- arbetsbok
wb = openpyxl.Workbook()
wsA = wb.active
wsA.title = 'Analys'
wsM = wb.create_sheet('Månad')
wsB = wb.create_sheet('Balans')
wsS = wb.create_sheet('Avstämning')
ws9 = wb.create_sheet('9140')
wsP = wb.create_sheet('Mappning')
wsR = wb.create_sheet('AARO 2609')
wsD = wb.create_sheet('Data')

# ---- Data
ND = len(data) + 1
hdr = ['Rad i CSV', 'Kostnadsställe', 'Konto', 'Underkonto', 'Kontosträng',
       'HB-datum', 'Månad', 'Belopp', 'Dok-typ', 'Dok-nr', 'Batchnr',
       'Batchdatum', 'Användar-id', 'Beskrivning', 'Förklaring', 'Söknamn',
       'Val-kod', 'Valutabelopp', 'AARO-rad']
wsD.append(hdr)
for c in wsD[1]:
    c.font = FET
for i, d in enumerate(data, start=2):
    wsD.append([d['rad'], d['mcu'], d['obj'], d['sub'],
                f'=B{i}&"."&C{i}&"."&D{i}', d['datum'], f'=MONTH(F{i})',
                d['belopp'], d['doktyp'], d['doknr'], d['batch'],
                d['batchdatum'], d['anv'], d['beskr'], d['forkl'],
                d['soknamn'], d['valkod'], d['valbelopp'] or None,
                f'=IFERROR(VLOOKUP(E{i},Mappning!$A:$C,3,0),"(ej mappad)")'])
    wsD.cell(i, 6).number_format = DATUM
    wsD.cell(i, 12).number_format = DATUM
    wsD.cell(i, 8).number_format = KR
    wsD.cell(i, 18).number_format = KR
wsD.freeze_panes = 'A2'
wsD.auto_filter.ref = f'A1:S{ND}'
for col, w in zip('ABCDEFGHIJKLMNOPQRS',
                  [9, 12, 7, 10, 18, 11, 7, 15, 7, 10, 10, 11, 11, 30, 24, 28, 7, 14, 11]):
    wsD.column_dimensions[col].width = w

# ---- Mappning
wsP.append(['Kontosträng', 'Beskrivning', 'AARO-rad', 'Källa', 'JDE 2609 Current'])
for c in wsP[1]:
    c.font = FET
for k in sorted(mapp):
    b, a, kalla, cur = mapp[k]
    wsP.append([k, b or konto_beskr.get(k, ''), a, kalla, cur])
    wsP.cell(wsP.max_row, 5).number_format = KR
wsP.freeze_panes = 'A2'
for col, w in zip('ABCDE', [18, 32, 11, 40, 17]):
    wsP.column_dimensions[col].width = w

# ---- AARO 2609 (kopia av Sheet1, indata)
s1hdr = [str(c.value).strip() for c in s1[1]]
wsR.append(s1hdr + ['Kontosträng'])
for c in wsR[1]:
    c.font = FET
for i, r in enumerate(aaro_rows, start=2):
    wsR.append(list(r) + [f'=H{i}&"."&J{i}&"."&L{i}'])
    for col in range(14, 19):
        wsR.cell(i, col).number_format = KR
NR = len(aaro_rows) + 1
wsR.freeze_panes = 'A2'
for col, w in zip('ABCDEFGHIJKLMNOPQRS',
                  [6, 9, 28, 8, 20, 11, 22, 9, 22, 6, 28, 6, 28, 16, 16, 10, 10, 10, 18]):
    wsR.column_dimensions[col].width = w

# ---- Månad (resultaträkning per AARO-rad och månad, tkr, resultattecken)
STRUKTUR = [
    ('rub', 'Nettoomsättning'), ('rad', '3010'),
    ('rub', 'Kostnad sålda varor'), ('rad', '4060'),
    ('sum', 'Bruttoresultat', 'BR'),
    ('kpi', 'Bruttomarginal', 'BR'),
    ('rub', 'Rörelsekostnader'),
    ('rad', '4511'), ('rad', '4531'), ('rad', '5013'), ('rad', '5740'),
    ('rad', '6360'), ('rad', '6361'), ('rad', '6363'), ('rad', 'MARKET10'),
    ('rad', 'MARKET40'), ('rad', 'MARKET60'), ('rad', 'ADMIN10'),
    ('sum', 'Summa rörelsekostnader', 'RK'),
    ('rub', 'Valutaeffekter och koncernpåslag'),
    ('rad', '7960'), ('rad', '7962'), ('rad', '799060'),
    ('sum', 'Rörelseresultat (EBIT)', 'EBIT'),
    ('kpi', 'Rörelsemarginal', 'EBIT'),
    ('rad', '8260'), ('rad', '8460'),
    ('sum', 'Resultat efter finansiella poster', 'EFN'),
    ('rad', '8860'),
    ('sum', 'Resultat före skatt', 'RFS'),
    ('rad', '8910'), ('rad', '8940'), ('rad', '8980'),
    ('sum', 'Periodens resultat', 'RES'),
]

wsM['A1'] = 'Electrolux Professional Sweden AB (PA), bolag 00047 – resultat per AARO-rad och månad 2026'
wsM['A1'].font = FET
wsM['A2'] = 'tkr, resultattecken (intäkt +, kostnad −) = −JDE-belopp/1000. Källa: flik Data (R55EL121 PA0001, uttag 2026-10-07 09:21:33).'
wsM['A2'].font = KURSIV
MH = 4
wsM.cell(MH, 1, 'AARO-rad').font = FET
wsM.cell(MH, 2, 'Benämning').font = FET
for m in range(1, MANAD + 1):
    c = wsM.cell(MH, 2 + m, m)
    c.font = FET
    c.number_format = '0'
wsM.cell(MH - 1, 3, 'Månad').font = FET
wsM.cell(MH, 3 + MANAD, f'Ack {MANADSNAMN[0]}–{MANADSNAMN[MANAD-1]}').font = FET
COL_ACK = 3 + MANAD
mrad = {}          # nyckel -> radnummer i Månad
radlista = []      # (radnr, typ)
r = MH + 1
grupp = []         # rader i aktuell summeringsgrupp sedan föregående summa
alla_sum = {}
for post in STRUKTUR:
    typ = post[0]
    if typ == 'rub':
        wsM.cell(r, 2, post[1]).font = FET
    elif typ == 'rad':
        kod = post[1]
        wsM.cell(r, 1, kod)
        wsM.cell(r, 2, beskr_aaro.get(kod, ''))
        for m in range(1, MANAD + 1):
            col = get_column_letter(2 + m)
            c = wsM.cell(r, 2 + m,
                         f'=-SUMIFS(Data!$H:$H,Data!$S:$S,$A{r},Data!$G:$G,{col}${MH})/1000')
            c.number_format = TKR
        c = wsM.cell(r, COL_ACK, f'=SUM(C{r}:{get_column_letter(2+MANAD)}{r})')
        c.number_format = TKR
        mrad[kod] = r
        grupp.append(r)
    elif typ == 'sum':
        namn, nyck = post[1], post[2]
        wsM.cell(r, 2, namn).font = FET
        for colnr in range(3, COL_ACK + 1):
            col = get_column_letter(colnr)
            if nyck == 'BR':
                f = f'={col}{mrad["3010"]}+{col}{mrad["4060"]}'
            elif nyck == 'RK':
                f = '=' + '+'.join(f'{col}{mrad[k]}' for k in
                                   ['4511', '4531', '5013', '5740', '6360', '6361', '6363',
                                    'MARKET10', 'MARKET40', 'MARKET60', 'ADMIN10'])
            elif nyck == 'EBIT':
                f = (f'={col}{alla_sum["BR"]}+{col}{alla_sum["RK"]}+{col}{mrad["7960"]}'
                     f'+{col}{mrad["7962"]}+{col}{mrad["799060"]}')
            elif nyck == 'EFN':
                f = f'={col}{alla_sum["EBIT"]}+{col}{mrad["8260"]}+{col}{mrad["8460"]}'
            elif nyck == 'RFS':
                f = f'={col}{alla_sum["EFN"]}+{col}{mrad["8860"]}'
            elif nyck == 'RES':
                f = f'={col}{alla_sum["RFS"]}+{col}{mrad["8910"]}+{col}{mrad["8940"]}+{col}{mrad["8980"]}'
            c = wsM.cell(r, colnr, f)
            c.number_format = TKR
            c.font = FET
        alla_sum[nyck] = r
        mrad[nyck] = r
    elif typ == 'kpi':
        namn, nyck = post[1], post[2]
        wsM.cell(r, 2, namn).font = KURSIV
        for colnr in range(3, COL_ACK + 1):
            col = get_column_letter(colnr)
            c = wsM.cell(r, colnr,
                         f'=IF({col}{mrad["3010"]}=0,0,{col}{alla_sum[nyck]}/{col}{mrad["3010"]})')
            c.number_format = PROC
            c.font = KURSIV
        mrad['KPI_' + nyck] = r
    r += 1
# kontroll mot hela resultatdelen i Data
r += 1
wsM.cell(r, 2, 'Kontroll: samtliga resultatkonton i Data (konto < 8000)').font = KURSIV
for m in range(1, MANAD + 1):
    col = get_column_letter(2 + m)
    c = wsM.cell(r, 2 + m, f'=-SUMIFS(Data!$H:$H,Data!$C:$C,"<8000",Data!$G:$G,{col}${MH})/1000')
    c.number_format = TKR
c = wsM.cell(r, COL_ACK, f'=SUM(C{r}:{get_column_letter(2+MANAD)}{r})')
c.number_format = TKR
mrad['KTRL_DATA'] = r
r += 1
wsM.cell(r, 2, 'Differens mot periodens resultat (ska vara 0)').font = KURSIV
for colnr in range(3, COL_ACK + 1):
    col = get_column_letter(colnr)
    c = wsM.cell(r, colnr, f'=ROUND({col}{mrad["KTRL_DATA"]}-{col}{mrad["RES"]},3)')
    c.number_format = TKR
mrad['KTRL_DIFF'] = r
wsM.freeze_panes = wsM.cell(MH + 1, 3)
wsM.column_dimensions['A'].width = 10
wsM.column_dimensions['B'].width = 36
for colnr in range(3, COL_ACK + 1):
    wsM.column_dimensions[get_column_letter(colnr)].width = 10
wsM.column_dimensions[get_column_letter(COL_ACK)].width = 11

# ---- Analys
SEP = get_column_letter(2 + MANAD)        # K
AUG = get_column_letter(1 + MANAD)        # J
ACK = get_column_letter(COL_ACK)          # L
wsA['A1'] = 'Electrolux Professional Sweden AB (PA), bolag 00047 – resultatanalys september 2026'
wsA['A1'].font = FET
wsA['A2'] = ('tkr, resultattecken (intäkt +, kostnad −). Källa: R55EL121 PA0001 uttag 2026-10-07 09:21:33 '
             '(flik Data) och Aaro_2609_inc_JDE_mapping.xlsx (flik AARO 2609). Utkast för granskning.')
wsA['A2'].font = KURSIV
AH = 4
rub = ['AARO-rad', 'Benämning', 'Sep 2026', 'Aug 2026', 'Snitt jan–aug', 'Sep mot snitt',
       'Ack jan–sep 2026', 'Andel av nettoomsättning ack']
for i, t in enumerate(rub, start=1):
    c = wsA.cell(AH, i, t)
    c.font = FET
    c.alignment = Alignment(wrap_text=True, vertical='bottom')
arad = {}
r = AH + 1
for post in STRUKTUR:
    typ = post[0]
    if typ == 'rub':
        wsA.cell(r, 2, post[1]).font = FET
        r += 1
        continue
    key = post[1] if typ == 'rad' else (post[2] if typ == 'sum' else 'KPI_' + post[2])
    mr = mrad[key]
    if typ == 'rad':
        wsA.cell(r, 1, f'=Månad!A{mr}')
        wsA.cell(r, 2, f'=Månad!B{mr}')
    else:
        wsA.cell(r, 2, f'=Månad!B{mr}')
    if typ == 'kpi':
        for colnr, src_col in [(3, SEP), (4, AUG), (7, ACK)]:
            c = wsA.cell(r, colnr, f'=Månad!{src_col}{mr}')
            c.number_format = PROC
            c.font = KURSIV
        c = wsA.cell(r, 5, f'=IF(E{arad["3010"]}=0,0,E{arad[post[2]]}/E{arad["3010"]})')
        c.number_format = PROC
        c.font = KURSIV
        wsA.cell(r, 2).font = KURSIV
    else:
        wsA.cell(r, 3, f'=Månad!{SEP}{mr}')
        wsA.cell(r, 4, f'=Månad!{AUG}{mr}')
        wsA.cell(r, 5, f'=AVERAGE(Månad!C{mr}:{AUG}{mr})')
        wsA.cell(r, 6, f'=C{r}-E{r}')
        wsA.cell(r, 7, f'=Månad!{ACK}{mr}')
        for colnr in range(3, 8):
            wsA.cell(r, colnr).number_format = TKR
        arad[key] = r
        if typ == 'sum':
            for colnr in range(2, 9):
                wsA.cell(r, colnr).font = FET
    if typ != 'kpi':
        c = wsA.cell(r, 8, f'=IF($G${arad["3010"]}=0,0,G{r}/$G${arad["3010"]})')
        c.number_format = PROC
    r += 1

# avstämning mot AARO-filen
r += 1
wsA.cell(r, 1, 'Avstämning av ackumulerat resultat mot AARO-filen').font = FET
r += 1
wsA.cell(r, 2, 'Periodens resultat ack enligt JDE-uttaget')
wsA.cell(r, 7, f'=G{arad["RES"]}').number_format = TKR
r_jde = r
r += 1
wsA.cell(r, 2, "Periodens resultat ack enligt AARO-filen (flik AARO 2609, BS/PL = PL, YTD 2026)")
wsA.cell(r, 7, "=-SUMIFS('AARO 2609'!$N:$N,'AARO 2609'!$A:$A,\"PL\")/1000").number_format = TKR
r_aaro = r
r += 1
wsA.cell(r, 2, 'Differens')
wsA.cell(r, 7, f'=G{r_jde}-G{r_aaro}').number_format = TKR
r_diff = r
r += 1
wsA.cell(r, 2, 'Förklaras av verifikat bokade 2026-10-06 efter AARO-uttaget:').font = KURSIV
r += 1
wsA.cell(r, 1, 'Dok-nr').font = FET
wsA.cell(r, 2, 'Text (Data, kolumn N)').font = FET
wsA.cell(r, 7, 'tkr').font = FET
r += 1
forkl_rader = []
texter = {26000344: 'JE Deftax garantireserv 2609 (konto 0473100.2911.000)',
          26000345: 'JE Koncernbidrag SE17 Q3 (konto 0473100.2806.000)'}
for v in SENA_VER:
    wsA.cell(r, 1, v)
    wsA.cell(r, 2, texter[v])
    wsA.cell(r, 7, f'=-SUMIFS(Data!$H:$H,Data!$J:$J,$A{r},Data!$I:$I,"JE",Data!$C:$C,"<8000")/1000').number_format = TKR
    forkl_rader.append(r)
    r += 1
wsA.cell(r, 2, 'Oförklarad differens (ska vara 0)')
wsA.cell(r, 7, f'=ROUND(G{r_diff}-SUM(G{forkl_rader[0]}:G{forkl_rader[-1]}),3)').number_format = TKR
r_ofk = r
r += 1
wsA.cell(r, 2, 'Rörelseresultat ack: JDE-uttaget mot AARO-filen (ska vara 0)')
ebit_aaro = ("-SUMPRODUCT(('AARO 2609'!$A$2:$A$%d=\"PL\")*('AARO 2609'!$B$2:$B$%d<>\"8260\")"
             "*('AARO 2609'!$B$2:$B$%d<>\"8460\")*('AARO 2609'!$B$2:$B$%d<>\"8860\")"
             "*('AARO 2609'!$B$2:$B$%d<>\"8910\")*('AARO 2609'!$B$2:$B$%d<>\"8940\")"
             "*('AARO 2609'!$B$2:$B$%d<>\"8980\"),'AARO 2609'!$N$2:$N$%d)/1000") % ((NR,) * 8)
wsA.cell(r, 7, f'=ROUND(G{arad["EBIT"]}-({ebit_aaro}),3)').number_format = TKR
r_ebitk = r
r += 2
wsA.cell(r, 1, 'Kontroller').font = FET
r += 1
wsA.cell(r, 2, 'Data: summa samtliga rader jan–sep (ska vara 0)')
wsA.cell(r, 7, '=ROUND(SUM(Data!$H:$H),2)').number_format = KR
r += 1
wsA.cell(r, 2, 'Data: resultatkonton utan AARO-rad, summa (ska vara 0)')
wsA.cell(r, 7, '=ROUND(SUMIFS(Data!$H:$H,Data!$S:$S,"(ej mappad)",Data!$C:$C,"<8000"),2)').number_format = KR
r += 1
wsA.cell(r, 2, 'Månad: differens mot samtliga resultatkonton, ack (ska vara 0)')
wsA.cell(r, 7, f'=Månad!{ACK}{mrad["KTRL_DIFF"]}').number_format = TKR

wsA.freeze_panes = wsA.cell(AH + 1, 3)
wsA.column_dimensions['A'].width = 10
wsA.column_dimensions['B'].width = 44
for col in 'CDEFG':
    wsA.column_dimensions[col].width = 12
wsA.column_dimensions['H'].width = 13
wsA.row_dimensions[AH].height = 30

# ---- Balans (tkr, JDE-tecken: debet +, kredit −)
wsB['A1'] = 'Electrolux Professional Sweden AB (PA), bolag 00047 – balansräkning per AARO-rad 2026-09-30'
wsB['A1'].font = FET
wsB['A2'] = ('tkr, JDE-tecken (debet +, kredit −). Utgående balans och septemberrörelse enligt AARO-filen; '
             'kolumn E visar verifikat i JDE-uttaget som bokats efter AARO-uttaget.')
wsB['A2'].font = KURSIV
BH = 4
for i, t in enumerate(['AARO-rad', 'Benämning', 'UB 2026-09-30 enligt AARO-fil',
                       'Rörelse sep enligt AARO-fil', 'Bokat efter AARO-uttaget (JE 26000344, 26000345)',
                       'Justerad UB 2026-09-30'], start=1):
    c = wsB.cell(BH, i, t)
    c.font = FET
    c.alignment = Alignment(wrap_text=True, vertical='bottom')
wsB.row_dimensions[BH].height = 45
bs_rader = []
seen = set()
for rr in aaro_rows:
    if rr[0] == 'BS' and rr[1] not in seen:
        seen.add(rr[1])
        bs_rader.append(rr[1])
r = BH + 1
wsB.cell(r, 2, 'Tillgångar').font = FET
r += 1
forsta = {}
def bs_rad(r, kod):
    wsB.cell(r, 1, kod)
    wsB.cell(r, 2, beskr_aaro.get(kod, ''))
    wsB.cell(r, 3, f"=SUMIFS('AARO 2609'!$N:$N,'AARO 2609'!$B:$B,$A{r})/1000")
    wsB.cell(r, 4, f"=SUMIFS('AARO 2609'!$O:$O,'AARO 2609'!$B:$B,$A{r})/1000")
    wsB.cell(r, 5, '=(' + '+'.join(
        f'SUMIFS(Data!$H:$H,Data!$S:$S,$A{r},Data!$J:$J,{v},Data!$I:$I,"JE")' for v in SENA_VER) + ')/1000')
    wsB.cell(r, 6, f'=C{r}+E{r}')
    for colnr in range(3, 7):
        wsB.cell(r, colnr).number_format = TKR
tillg = [k for k in bs_rader if k.startswith('1')]
skuld = [k for k in bs_rader if k.startswith('2')]
r0 = r
for k in tillg:
    bs_rad(r, k)
    r += 1
wsB.cell(r, 2, 'Summa tillgångar').font = FET
for colnr in range(3, 7):
    col = get_column_letter(colnr)
    c = wsB.cell(r, colnr, f'=SUM({col}{r0}:{col}{r-1})')
    c.number_format = TKR
    c.font = FET
r_st = r
r += 2
wsB.cell(r, 2, 'Eget kapital och skulder').font = FET
r += 1
r1 = r
for k in skuld:
    bs_rad(r, k)
    r += 1
wsB.cell(r, 2, 'Summa eget kapital och skulder').font = FET
for colnr in range(3, 7):
    col = get_column_letter(colnr)
    c = wsB.cell(r, colnr, f'=SUM({col}{r1}:{col}{r-1})')
    c.number_format = TKR
    c.font = FET
r_sk = r
r += 2
wsB.cell(r, 2, 'Netto balansräkning (debet + motsvarar periodens vinst)')
for colnr in (3, 5, 6):
    col = get_column_letter(colnr)
    wsB.cell(r, colnr, f'={col}{r_st}+{col}{r_sk}').number_format = TKR
r_net = r
r += 1
wsB.cell(r, 2, 'Kontroll: netto + periodens resultat enligt Analys (ska vara 0)').font = KURSIV
wsB.cell(r, 6, f'=ROUND(F{r_net}-Analys!G{arad["RES"]},3)').number_format = TKR
wsB.cell(r, 3, f'=ROUND(C{r_net}-Analys!G{r_aaro},3)').number_format = TKR
wsB.freeze_panes = wsB.cell(BH + 1, 3)
wsB.column_dimensions['A'].width = 10
wsB.column_dimensions['B'].width = 34
for col in 'CDEF':
    wsB.column_dimensions[col].width = 14

# ---- Avstämning per konto (kr)
wsS['A1'] = 'Avstämning per JDE-konto: JDE-uttaget (flik Data) mot AARO-filen (flik AARO 2609) och JDE 2609-fliken'
wsS['A1'].font = FET
wsS['A2'] = ('kr, JDE-tecken. Uttaget innehåller endast 2026 års transaktioner; för balanskonton jämförs därför '
             'septemberrörelsen, och kolumn M visar härledd ingående balans 2026-01-01 (AARO UB − uttagets rörelse '
             'exklusive verifikat bokade efter AARO-uttaget).')
wsS['A2'].font = KURSIV
SH = 4
for i, t in enumerate(['Kontosträng', 'Beskrivning', 'RR/BR', 'AARO-rad', 'Uttaget ack jan–sep',
                       'Uttaget sep', 'Varav bokat efter AARO-uttaget (JE 26000344, 26000345)',
                       'AARO-fil YTD/UB', 'AARO-fil månad', 'JDE 2609-flik Current',
                       'Diff sep (F−G−I)', 'Diff ack RR (E−G−H)', 'Härledd IB 2026-01-01 BR (H−(E−G))',
                       'Förklaring'], start=1):
    c = wsS.cell(SH, i, t)
    c.font = FET
    c.alignment = Alignment(wrap_text=True, vertical='bottom')
wsS.row_dimensions[SH].height = 45
konton = sorted(set(mapp) | {d['konto'] for d in data} |
                {f'{rr[7]}.{rr[9]}.{rr[11]}' for rr in aaro_rows})
forkl = {
    '0473100.2806.000': 'JE 26000345 Koncernbidrag SE17 Q3, batchdatum 2026-10-06, ej med i AARO-filen',
    '0470000.9107.000': 'JE 26000345 Koncernbidrag SE17 Q3, batchdatum 2026-10-06, ej med i AARO-filen',
    '0473100.2911.000': 'JE 26000344 Deftax garantireserv 2609, batchdatum 2026-10-06, ej med i AARO-filen',
    '0470000.8220.001': 'JE 26000344 Deftax garantireserv 2609, batchdatum 2026-10-06, ej med i AARO-filen',
    '0470000.8401.SPS': 'Finns i AARO-filen (AARO 1440) men saknas i JDE 2609-fliken, som därför inte summerar till noll',
}
r = SH + 1
for k in konton:
    obj = int(k.split('.')[1])
    wsS.cell(r, 1, k)
    wsS.cell(r, 2, f'=IFERROR(VLOOKUP(A{r},Mappning!$A:$B,2,0),"")')
    wsS.cell(r, 3, 'RR' if obj < 8000 else 'BR')
    wsS.cell(r, 4, f'=IFERROR(VLOOKUP(A{r},Mappning!$A:$C,3,0),"(ej mappad)")')
    wsS.cell(r, 5, f'=SUMIFS(Data!$H:$H,Data!$E:$E,$A{r})')
    wsS.cell(r, 6, f'=SUMIFS(Data!$H:$H,Data!$E:$E,$A{r},Data!$G:$G,{MANAD})')
    wsS.cell(r, 7, '=' + '+'.join(
        f'SUMIFS(Data!$H:$H,Data!$E:$E,$A{r},Data!$J:$J,{v},Data!$I:$I,"JE")' for v in SENA_VER))
    wsS.cell(r, 8, f"=SUMIFS('AARO 2609'!$N:$N,'AARO 2609'!$S:$S,$A{r})")
    wsS.cell(r, 9, f"=SUMIFS('AARO 2609'!$O:$O,'AARO 2609'!$S:$S,$A{r})")
    wsS.cell(r, 10, f'=SUMIFS(Mappning!$E:$E,Mappning!$A:$A,$A{r})')
    wsS.cell(r, 11, f'=ROUND(F{r}-G{r}-I{r},2)')
    wsS.cell(r, 12, f'=IF(C{r}="RR",ROUND(E{r}-G{r}-H{r},2),"")')
    wsS.cell(r, 13, f'=IF(C{r}="BR",ROUND(H{r}-(E{r}-G{r}),2),"")')
    wsS.cell(r, 14, forkl.get(k, ''))
    for colnr in range(5, 14):
        wsS.cell(r, colnr).number_format = KR
    r += 1
r_s_end = r - 1
r += 1
wsS.cell(r, 2, 'Summa').font = FET
for colnr in range(5, 14):
    col = get_column_letter(colnr)
    c = wsS.cell(r, colnr, f'=SUM({col}{SH+1}:{col}{r_s_end})')
    c.number_format = KR
    c.font = FET
wsS.freeze_panes = wsS.cell(SH + 1, 2)
wsS.auto_filter.ref = f'A{SH}:N{r_s_end}'
for col, w in zip('ABCDEFGHIJKLMN', [18, 30, 6, 10, 16, 16, 16, 16, 16, 16, 14, 14, 16, 60]):
    wsS.column_dimensions[col].width = w

# ---- 9140.000
K9 = '0470000.9140.000'
ws9['A1'] = 'Specifikation konto 0470000.9140.000 ACCRUED EXP./PREPAID INC.,EXT. (AARO 291010) per 2026-09-30'
ws9['A1'].font = FET
ws9['A2'] = ('kr, JDE-tecken (kredit − = upplupen kostnad). Transaktionerna hämtas med cellreferens från flik Data; '
             'kolumnerna I (Post) och J (Motkonto enligt verifikatet) är bedömningar/indata.')
ws9['A2'].font = KURSIV
H9 = 4
for i, t in enumerate(['Rad i Data', 'HB-datum', 'Dok-typ', 'Dok-nr', 'Batchdatum', 'Användar-id',
                       'Text', 'Belopp', 'Post', 'Motkonto enligt verifikatet'], start=1):
    ws9.cell(H9, i, t).font = FET
mot = {}
for d in data:
    key = (d['doktyp'], d['doknr'], d['datum'])
    if d['konto'] != K9:
        mot.setdefault(key, []).append(d['konto'])
def klassa(d):
    t = (d['beskr'] + ' ' + d['soknamn']).lower()
    if 'revision' in t:
        return 'Revision'
    if 'fora' in t:
        return 'Fora'
    if 'credendo' in t:
        return 'Credendo'
    if 'getronics' in t:
        return 'Getronics'
    if 'one connect' in t:
        return 'One Connected'
    if 'tempestive' in t:
        return 'Tempestive'
    return 'Övrigt'
r = H9 + 1
rad9 = []
for i, d in enumerate(data, start=2):
    if d['konto'] != K9:
        continue
    ws9.cell(r, 1, i)
    ws9.cell(r, 2, f'=Data!F{i}').number_format = DATUM
    ws9.cell(r, 3, f'=Data!I{i}')
    ws9.cell(r, 4, f'=Data!J{i}')
    ws9.cell(r, 5, f'=Data!L{i}').number_format = DATUM
    ws9.cell(r, 6, f'=Data!M{i}')
    ws9.cell(r, 7, f'=Data!N{i}')
    ws9.cell(r, 8, f'=Data!H{i}').number_format = KR
    post = klassa(d)
    # Tempestive och One Connected i juni–juli: 9140-delen avser One Connected (230 000/mån),
    # Tempestive-delen gick mot 8240.000
    if 'one connect' in d['beskr'].lower() and 'tempestive' in d['beskr'].lower():
        post = 'One Connected'
    ws9.cell(r, 9, post)
    ws9.cell(r, 10, ', '.join(sorted(set(mot.get((d['doktyp'], d['doknr'], d['datum']), [])))) or
             'Leverantörsfaktura (V-batch), ingen motrad i uttaget')
    rad9.append((r, d, post))
    r += 1
r9_end = r - 1
ws9.cell(r, 7, 'Summa rörelse 2026').font = FET
c = ws9.cell(r, 8, f'=SUM(H{H9+1}:H{r9_end})')
c.number_format = KR
c.font = FET
r_9sum = r
r += 1
ws9.cell(r, 7, 'Kontroll mot Data (ska vara 0)').font = KURSIV
ws9.cell(r, 8, f'=ROUND(H{r_9sum}-SUMIFS(Data!$H:$H,Data!$E:$E,"{K9}"),2)').number_format = KR
r += 2
ws9.cell(r, 1, 'Sammanställning per post').font = FET
r += 1
for i, t in enumerate(['Post', '', '', 'Härledd IB 2026-01-01', 'Rörelse 2026', 'UB 2026-09-30',
                       'Underlag för härledd IB'], start=1):
    if t:
        ws9.cell(r, i, t).font = FET
r += 1
poster = ['One Connected', 'Credendo', 'Revision', 'Getronics', 'Tempestive', 'Fora']
ps = {}
cred_tidiga = [rr for rr, d, p in rad9 if p == 'Credendo' and d['datum'] <= datetime(2026, 3, 31)]
get_rad = [rr for rr, d, p in rad9 if p == 'Getronics']
for p in poster:
    ws9.cell(r, 1, p)
    if p == 'Credendo':
        ws9.cell(r, 4, '=-(' + '+'.join(f'H{x}' for x in cred_tidiga) + ')')
        ws9.cell(r, 7, 'Fakturor 2026-03-01 bokade mot kontot minus JE 26000101 (överskjutande del till kostnad)')
    elif p == 'Getronics':
        ws9.cell(r, 4, '=-(' + '+'.join(f'H{x}' for x in get_rad) + ')')
        ws9.cell(r, 7, 'Upplösning JE 26000323 2026-09-30 utan reservering under 2026')
    else:
        ws9.cell(r, 4, 0)
        ws9.cell(r, 7, 'Ingen IB antagen; första bokning 2026' if p != 'Fora' else
                 'Fakturor konterade på kontot, ombokade JE 26000267')
    ws9.cell(r, 5, f'=SUMIFS($H${H9+1}:$H${r9_end},$I${H9+1}:$I${r9_end},$A{r})')
    ws9.cell(r, 6, f'=D{r}+E{r}')
    for colnr in (4, 5, 6):
        ws9.cell(r, colnr).number_format = KR
    ps[p] = r
    r += 1
ws9.cell(r, 1, 'Ej specificerad del av IB (kräver specifikation 2025-12-31)')
ws9.cell(r, 4, f'=D{r+2}-SUM(D{ps[poster[0]]}:D{r-1})')
ws9.cell(r, 5, 0)
ws9.cell(r, 6, f'=D{r}+E{r}')
for colnr in (4, 5, 6):
    ws9.cell(r, colnr).number_format = KR
r_ospec = r
r += 1
ws9.cell(r, 1, 'Summa').font = FET
for colnr in (4, 5, 6):
    col = get_column_letter(colnr)
    c = ws9.cell(r, colnr, f'=SUM({col}{ps[poster[0]]}:{col}{r-1})')
    c.number_format = KR
    c.font = FET
r_9tot = r
r += 1
ws9.cell(r, 1, 'IB härledd som UB enligt AARO-filen minus rörelse 2026')
ws9.cell(r, 4, f'=F{r+1}-E{r}').number_format = KR
ws9.cell(r, 5, f'=H{r_9sum}').number_format = KR
r_ib = r
r += 1
ws9.cell(r, 1, 'UB enligt AARO-filen (flik AARO 2609, YTD 2026)')
ws9.cell(r, 6, f"=SUMIFS('AARO 2609'!$N:$N,'AARO 2609'!$S:$S,\"{K9}\")").number_format = KR
r += 1
ws9.cell(r, 1, 'Kontroll: summa UB mot AARO-filen (ska vara 0)').font = KURSIV
ws9.cell(r, 6, f'=ROUND(F{r_9tot}-F{r-1},2)').number_format = KR
# rättning: raden för ej specificerad IB pekar på härledd IB-raden
ws9.cell(r_ospec, 4, f'=D{r_ib}-SUM(D{ps[poster[0]]}:D{r_ospec-1})')
r += 2
ws9.cell(r, 1, 'Revisionsreserv: bokat ack mot 12 000 kr per månad').font = FET
r += 1
ws9.cell(r, 1, 'Bokat ack 2026 (JE Reserv revision)')
ws9.cell(r, 6, f'=-E{ps["Revision"]}').number_format = KR
r_rb = r
r += 1
ws9.cell(r, 1, 'Månadsbelopp revisionsreserv (indata)')
ws9.cell(r, 6, 12000).number_format = KR
r_mb = r
r += 1
ws9.cell(r, 1, 'Antal månader jan–sep')
ws9.cell(r, 6, MANAD)
r_am = r
r += 1
ws9.cell(r, 1, 'Förväntat ack')
ws9.cell(r, 6, f'=F{r_mb}*F{r_am}').number_format = KR
r_fa = r
r += 1
ws9.cell(r, 1, 'Differens bokat mot förväntat')
ws9.cell(r, 6, f'=F{r_rb}-F{r_fa}').number_format = KR
ws9.freeze_panes = ws9.cell(H9 + 1, 1)
for col, w in zip('ABCDEFGHIJ', [11, 11, 8, 10, 11, 11, 28, 15, 22, 44]):
    ws9.column_dimensions[col].width = w

wb.save(UT)
print('sparad', UT, 'datarader', len(data), 'mappning', len(mapp))
