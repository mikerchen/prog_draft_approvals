import pandas as pd
from datetime import datetime

#Import Approvals and Prisma Workbook
approvals = pd.read_excel('Input/Approvals Grid.xlsx', sheet_name = 0)
prisma = pd.read_excel('Input/Prisma.xlsx', header=1, sheet_name = 0)
# fees_sums = pd.read_excel('summed sheet')

print('Rows found in Draft Approvals Sheet: ' + str(approvals['Est'].count()))
print('Rows found in Prisma Sheet: ' + str(prisma['Estimate code'].count()))

prisma = prisma.rename(columns={'Month of service': 'Month of Service', 'Estimate code':'Est', 'Supplier short name': 'Publisher','Billed':'Ordered'})

#Convert Month of Service to datetime
approvals['Month of Service'] = pd.to_datetime(approvals['Month of Service'])
prisma['Month of Service'] = pd.to_datetime(prisma['Month of Service'])

match = approvals.merge(prisma, how='inner',left_on=['Est','Month of Service','Publisher'],right_on=['Est','Month of Service','Publisher'],indicator=True)

match['Approval Status'] = ''
match['Estimate Status'] = ''

for i, j in match.iterrows():
    if(match['Product Name'][i] != 'Fees'):
        if(match['Ordered'][i] == match['Actual Net Billable'][i]):
            match.at[i,'Approval Status'] = 'Media Buy: Approved'
        else:
            diff = match['Ordered'][i] - match['Actual Net Billable'][i]
            diff_cur = "${:,.2f}".format(diff)
            match.at[i,'Approval Status'] = 'Media Buy (Ordered <> Billable): Delta = ' + str(diff_cur)
    else:
        match.at[i, 'Approval Status'] = 'Fee Buy: Unknown'

for i, j in match.iterrows():
    if('Media Buy: Ordered <> Billable' in match['Approval Status'][i]):
        estimate = match['Est'][i]
        for k, w in match.iterrows():
            if(match['Est'][k] == estimate):
                match.at[k,'Estimate Status'] = 'Dependent'

# fees_sums

# Estimate Key against a summed sheet
# 	If fee is reconciled or fee does not have a previous month's billing, check Ordered vs Billed
# 	If fee isn't reconciled, then fee must be within $300 +/-, previous month OR a delta (Calculated Fee - previous month's ordered)

match.to_csv('test.csv')