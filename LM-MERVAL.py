from datetime import date,timedelta
from pyhomebroker import HomeBroker
import pandas as pd
from openpyxl import Workbook

################################################## IMPORTANTE ##################################################
# Si no encotras el archivo .xlsx generado , buscalo en tu disco local > user o usuarios > nombre de tu usuario ej:
# C:\Users\pepe

# Credenciales de tu cuenta HomeBroker
username = ''
password = ''
dni = ''
broker = 000 #VERIFICAR SIEMPRE NÚMERO DE TU BROKER

end_date = date.today()
start_date = end_date - timedelta(days=365)

# Crear instancia de HomeBroker y autenticar
hb = HomeBroker(int(broker))
hb.auth.login(dni=dni, user=username, password=password, raise_exception=True)

# Código del mercado
TICKETS = [
    "ALUA", "BBAR", "BMA", "BYMA", "CEPU", "CRES", "COME", "EDN",
    "GGAL", "LOMA", "MIRG", "PAMP", "SUPV", "TECO2", "TGNO4", "TGSU2",
    "TRAN", "TXAR", "VALO", "YPFD"
]

GALPONES = [
    "AGRO", "AUSO", "BHIP", "BOLT", "BPAT",
    "CADO", "CAPX", "CARC", "CECO2", "CELU", "CGPA2", "CRE3W", "CTIO", "CVH", "DGCE", "DGCU2", "DOME", "FERR",
    "FIPL", "GAMI", "GBAN", "GCDI", "GCLA", "HARG", "HAVA", "HSAT", "INTR", "INVJ", "LEDE", "LONG", "METR",
    "MOLA", "MOLI", "MORI", "MTR", "OEST", "PATA", "RICH", "RIGO", "ROSE", "SAMI", "SEMI"
]

df_merval = pd.DataFrame()
results = pd.DataFrame()

df_galpones = pd.DataFrame()
results_galpones = pd.DataFrame()

for ticket in TICKETS:
    data = hb.history.get_daily_history(ticket, start_date, end_date)
    df_merval[ticket] = data['close']
    highValue = data['high'].max()
    maxValueDate = data['date'].iloc[data['high'].idxmax()]
    lastValue = df_merval[ticket].iloc[-1]
    percentDif = ((lastValue - highValue) / highValue) * 100
    nameTicket = ticket.split(".")[0]
    results[ticket] = nameTicket, lastValue, highValue, maxValueDate, percentDif

for ticket in GALPONES:
    data = hb.history.get_daily_history(ticket, start_date, end_date)
    df_galpones[ticket] = data['close']
    highValue = data['high'].max()
    maxValueDate = data['date'].iloc[data['high'].idxmax()]
    lastValue = df_galpones[ticket].iloc[-1]
    percentDif = ((lastValue - highValue) / highValue) * 100
    nameTicket = ticket.split(".")[0]
    results_galpones[ticket] = nameTicket, lastValue, highValue, maxValueDate, percentDif

results = results.T
results.columns = ['Acción', 'Ultimo Precio', 'Maximo', 'Fecha Max', '%']
results.sort_values(by='%', ascending=False, inplace=True)

results_galpones = results_galpones.T
results_galpones.columns = ['Acción', 'Ultimo Precio', 'Maximo', 'Fecha Max', '%']
results_galpones.sort_values(by='%', ascending=False, inplace=True)

writer = pd.ExcelWriter('BRECHA1.xlsx', engine='openpyxl')
results.to_excel(writer, index=False, sheet_name='Panel Lider',startrow=1, startcol=1)
results_galpones.to_excel(writer, index=False, sheet_name='Panel General',startrow=1, startcol=1)


workbook = writer.book
worksheet = writer.sheets['Panel Lider']
worksheet = writer.sheets['Panel General']

# Guardar el libro de trabajo usando openpyxl
workbook.save('BRECHA1.xlsx')