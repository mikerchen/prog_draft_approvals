import pandas as pd
from datetime import date
import os

now = date.today().strftime('%B_%d_%Y')
path = os.path.join('Output/','Fee Output_' + now + '/')

if (os.path.isdir(path)):
    print(path + ' directory exists')
else:
    os.mkdir(path)
    print(path  + ' directory created')

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
dedup_cols = ['Estimate code', 'Month', 'Publisher', 'Campaign']
prisma = prisma.drop_duplicates(subset=dedup_cols, keep='first')

prisma.to_csv(path + 'Prisma_deduped_' + now + '.csv')

print('Calculating DCM fees...')
for i, j in dcm.iterrows():
    if dcm['Creative Type'][i] in dcm_rates.keys():
        fee = (dcm['Impressions'][i]/1000)*dcm_rates[dcm['Creative Type'][i]]
        fee_cur = "{:.2f}".format(fee)
        dcm.at[i,'Cost'] = fee_cur
    else:
        dcm.at[i,'Cost'] = 'Invalid Creative Type: ' + str(dcm['Creative Type'][i])
    print('DCM Fees: ' + str(i+1) + '/' + str(dcm['Creative Type'].count()) + ' Complete')

dcm_joined = dcm.merge(prisma, how='left', on=['Month', 'Publisher', 'Campaign'] )
print('Prisma matched to DCM: ' + str(dcm_joined['Estimate code'].count()) + ' matches found.')

dcm_joined.to_csv(path + 'DCM_Fees_' + now + '.csv')

dcm_joined['Cost'] = pd.to_numeric(dcm_joined['Cost'])

dcm_grouped = dcm_joined.groupby(['Estimate code','Month','Publisher'])['Cost'].sum()
dcm_grouped.to_csv(path + 'DCM_Grouped_' + now + '.csv')

clinch = pd.read_excel('Input/Clinch Report.xlsx', sheet_name = 0)


# clinch, DCM, Invoice report

# Match
# Ingest invoice file, map reconciled status
# calculate DCM
# plot against keys

# Output: Summed Sheet