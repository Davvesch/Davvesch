"""Specifikation av utestående belopp på 0470000.9140.000 (PA, bolag 00047).
Anrop: python3 -I spec9140.py <csv 2025> <csv 2026> <ut.xlsx>"""
import sys
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment
import os; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from load import load

K = '0470000.9140.000'
FET, KURS = Font(bold=True), Font(italic=True)
KR = '#,##0;-#,##0;0'
filer = [(sys.argv[1], 'R55EL121_PA0001_24651425 (2025)'), (sys.argv[2], 'R55EL121_PA0001_24648959 (2026)')]

def post(t):
    t = t.lower()
    for nyck, namn in [('revision', 'Revision'), ('fora', 'Fora'), ('credendo', 'Credendo'),
                       ('getronics', 'Getronics'), ('one connect', 'One Connected'),
                       ('tempest', 'Tempestive'), ('löneskatt', 'Löneskatt 2025'),
                       ('slutskatt', 'Slutskatt 2024')]:
        if nyck in t:
            return namn
    return 'Vända reserveringar okt 2025'

rader = []
for p, namn in filer:
    df = load(p)
    df['Belopp_n'] = df['Belopp_n'].fillna(0)
    x = df[df['acct'] == K]
    for _, r in x.iterrows():
        v = df[(df['Dok- typ'] == r['Dok- typ']) & (df['Dokument- nummer'] == r['Dokument- nummer'])
               & (df['HB- datum'] == r['HB- datum']) & (df['acct'] != K)]
        mot = ', '.join(sorted(set(v['acct']))) or 'Leverantörsfaktura'
        ps = post(r['Beskrivning söknamn'] + ' ' + r['Söknamn'])
        if 'one connect' in r['Beskrivning söknamn'].lower():
            ps = 'One Connected'
        rader.append([namn, r['rad'], r['datum'].to_pydatetime(), r['Dok- typ'], int(r['Dokument- nummer']),
                      datetime.strptime(r['Batch- datum'], '%y-%m-%d'), r['Användar- id'],
                      r['Beskrivning söknamn'], mot, float(r['Belopp_n']), ps])

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Spec 9140'
wt = wb.create_sheet('Transaktioner')

# Transaktioner
hdr = ['Källa', 'Rad i källfil', 'HB-datum', 'År', 'Dok-typ', 'Dok-nr', 'Batchdatum', 'Användar-id',
       'Text', 'Motkonto enligt verifikatet', 'Belopp', 'Post']
wt.append(hdr)
for c in wt[1]:
    c.font = FET
# härledd ingående balans 2025-01-01
wt.append(['Härledd: UB 2026-09-30 enligt JDE 2609-fliken minus rörelser 2025–2026 (se Spec 9140)', None,
           datetime(2025, 1, 1), '=YEAR(C2)', 'IB', None, None, None,
           'Ingående balans 2025-01-01', None, "='Spec 9140'!C25", 'Slutskatt 2024'])
for r in rader:
    wt.append(r[:3] + [None] + r[3:])
for i in range(2, wt.max_row + 1):
    if i > 2:
        wt.cell(i, 4, f'=YEAR(C{i})')
    wt.cell(i, 3).number_format = 'yyyy-mm-dd'
    wt.cell(i, 7).number_format = 'yyyy-mm-dd'
    wt.cell(i, 11).number_format = KR
NT = wt.max_row
wt.cell(NT + 1, 9, 'Summa').font = FET
c = wt.cell(NT + 1, 11, f'=SUM(K2:K{NT})'); c.number_format = KR; c.font = FET
wt.freeze_panes = 'A2'
wt.auto_filter.ref = f'A1:L{NT}'
for col, w in zip('ABCDEFGHIJKL', [32, 11, 11, 6, 8, 10, 11, 11, 26, 44, 13, 26]):
    wt.column_dimensions[col].width = w

# Spec
ws['A1'] = 'Electrolux Professional Sweden AB (PA), bolag 00047 – specifikation konto 0470000.9140.000 per 2026-09-30'
ws['A1'].font = FET
ws['A2'] = ('kr, JDE-tecken (kredit − = upplupen kostnad). Underlag: R55EL121 PA0001 för 2025 och jan–sep 2026, '
            'flik Transaktioner. Utkast.')
ws['A2'].font = KURS
H = 4
for i, t in enumerate(['Post', 'Bokad av', 'IB 2025-01-01', 'Rörelse 2025', 'Rörelse 2026',
                       'Utestående 2026-09-30', 'Kommentar'], start=1):
    c = ws.cell(H, i, t); c.font = FET; c.alignment = Alignment(wrap_text=True, vertical='bottom')
poster = [
    ('One Connected', 'David Scherling, Marcus Kullenberg',
     'Reserv 230 000 per månad februari–juli 2026 (fem månader); 920 000 upplöst 2026-09-30, JE 26000321.'),
    ('Credendo', 'David Scherling',
     'Reserv 2025-12-31 om 508 550 (JE 25000072) reglerad mot fakturor i mars 2026; ny reserv per 2026-09-30, JE 26000332.'),
    ('Revision', 'Anna-Karin Nelson',
     '60 000 i juni och 12 000 per månad juli–september 2026, totalt 96 000 (åtta månader).'),
    ('Getronics', 'Marcus Kullenberg, David Scherling',
     'Reserverat 71 000 i oktober och november 2025 (JE 25000025, 25000049) och 669 500 i december 2025 (JE 25000071); '
     '669 500 upplöst 2026-09-30 (JE 26000323). Kvarstående 142 000 avser oktober–november 2025.'),
    ('Löneskatt 2025', 'Anna-Karin Nelson', 'UO 109158 per 2025-12-31.'),
    ('Tempestive', 'David Scherling, Marcus Kullenberg',
     'Reserverat oktober–november 2025, vänt december 2025 (JE 25000076); reserverat januari–mars 2026 och upplöst april 2026.'),
    ('Slutskatt 2024', 'Anna-Karin Nelson', 'Ingående balans 2025 upplöst 2025-12-31, UO 109157.'),
    ('Fora', 'Leverantörsfakturor, ombokade av Anna-Karin Nelson',
     'Fyra fakturor 2026 konterade direkt på kontot, ombokade 2026-08-31 (JE 26000267).'),
    ('Vända reserveringar okt 2025', 'David Scherling', 'JE 25000010 och 25000018, vända 2025-11-01.'),
]
r = H + 1
for p, av, kom in poster:
    ws.cell(r, 1, p); ws.cell(r, 2, av); ws.cell(r, 7, kom)
    ws.cell(r, 3, f'=SUMIFS(Transaktioner!$K:$K,Transaktioner!$L:$L,$A{r},Transaktioner!$E:$E,"IB")')
    ws.cell(r, 4, f'=SUMIFS(Transaktioner!$K:$K,Transaktioner!$L:$L,$A{r},Transaktioner!$D:$D,2025,Transaktioner!$E:$E,"<>IB")')
    ws.cell(r, 5, f'=SUMIFS(Transaktioner!$K:$K,Transaktioner!$L:$L,$A{r},Transaktioner!$D:$D,2026)')
    ws.cell(r, 6, f'=C{r}+D{r}+E{r}')
    for col in range(3, 7):
        ws.cell(r, col).number_format = KR
    ws.cell(r, 7).alignment = Alignment(wrap_text=True, vertical='top')
    for col in range(1, 7):
        ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical='top')
    r += 1
last = r - 1
ws.cell(r, 1, 'Summa').font = FET
for col in range(3, 7):
    L = 'CDEF'[col - 3]
    c = ws.cell(r, col, f'=SUM({L}{H+1}:{L}{last})'); c.number_format = KR; c.font = FET
rs = r
r += 2
ws.cell(r, 1, 'Avstämning').font = FET
r += 1
ws.cell(r, 1, 'UB 2026-09-30 enligt JDE 2609-fliken (Aaro_2609_inc_JDE_mapping.xlsx, rad 203)')
ws.cell(r, 6, -1041475).number_format = KR
r_ub = r
r += 1
ws.cell(r, 1, 'Rörelse 2025 och 2026 enligt uttagen')
ws.cell(r, 6, f'=D{rs}+E{rs}').number_format = KR
r_ror = r
r += 1
ws.cell(r, 1, 'Härledd IB 2025-01-01 (förs till flik Transaktioner, rad 2)')
ws.cell(r, 3, f'=F{r_ub}-F{r_ror}').number_format = KR
wt.cell(2, 11, f"='Spec 9140'!C{r}")
r += 1
ws.cell(r, 1, 'Kontroll: summa utestående mot UB (ska vara 0)').font = KURS
ws.cell(r, 6, f'=ROUND(F{rs}-F{r_ub},2)').number_format = KR
ws.freeze_panes = 'A5'
for col, w in zip('ABCDEFG', [28, 24, 13, 13, 13, 15, 70]):
    ws.column_dimensions[col].width = w
ws.row_dimensions[H].height = 30
wb.save(sys.argv[3])
print('ok', len(rader))
