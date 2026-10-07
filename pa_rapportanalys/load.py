import csv, pandas as pd
def load(p):
    with open(p, encoding='utf-8-sig', newline='') as f:
        rows = list(csv.reader(f, delimiter=';'))
    hdr = [h.strip().replace('\xa0','') for h in rows[4]]
    data = [r + ['']*(len(hdr)-len(r)) for r in rows[5:]]
    df = pd.DataFrame([r[:len(hdr)] for r in data], columns=hdr)
    df.insert(0,'rad', range(6, 6+len(df)))
    df['Belopp_n'] = pd.to_numeric(df['Belopp'].str.replace(' ','').str.replace(',','.'), errors='coerce')
    df['mcu'] = df['Centro Costo'].str.lstrip('_')
    df['obj'] = df['Konto']
    df['sub'] = df['Und- kto'].str.lstrip('_')
    df['acct'] = df['mcu']+'.'+df['obj']+'.'+df['sub']
    df['datum'] = pd.to_datetime(df['HB- datum'], format='%y-%m-%d', errors='coerce')
    df['mån'] = df['datum'].dt.month
    return df
