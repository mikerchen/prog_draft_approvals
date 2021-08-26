from numpy import int64
import pandas as pd
from datetime import date

now = date.today().strftime('%B_%d_%Y')

dcm_rates = {
    'Display': 0.035,
    'Tracking': 0.035,
    'In-Stream Audio': 0.250,
    'In-Stream Video': 0.250,
    'Rich Media display banner': 0.110
}

dcm = pd.read_excel('Input/DCM.xlsx', sheet_name = 0)
prisma = pd.read_excel('Input/Prisma.xlsx', header=1, sheet_name = 0, usecols=['Campaign name','Estimate code', 'Month of service', 'Supplier short name'])

dcm['Month'] = pd.to_datetime(dcm['Month'])
dcm['Publisher'] = 'DCM'
dcm['Cost'] = ''

prisma = prisma.rename(columns={'Month of service': 'Month', 'Supplier short name': 'Publisher', 'Campaign name': 'Campaign'})
prisma['Month'] = pd.to_datetime(prisma['Month'])

for i, j in dcm.iterrows():
    if dcm['Creative Type'][i] in dcm_rates.keys():
        fee = (dcm['Impressions'][i]/1000)*dcm_rates[dcm['Creative Type'][i]]
        fee_cur = "{:.2f}".format(fee)
        dcm.at[i,'Cost'] = fee_cur
    else:
        dcm.at[i,'Cost'] = 'Invalid Creative Type: ' + str(dcm['Creative Type'][i])

dcm_joined = dcm.merge(prisma, how='left', on=['Month', 'Publisher', 'Campaign'] )

dcm_joined.to_csv('Output/DCM_Fees_' + now + '.csv')

dcm_joined['Cost'] = pd.to_numeric(dcm_joined['Cost'])

dcm_grouped = dcm_joined.groupby(['Estimate code','Month','Publisher'])['Cost'].sum()
dcm_grouped.to_csv('Output/DCM_Grouped_' + now + '.csv')


print('Process Successful')

# clinch, DCM, Invoice report

# Match
# Ingest invoice file, map reconciled status
# calculate DCM
# plot against keys

# Output: Summed Sheet