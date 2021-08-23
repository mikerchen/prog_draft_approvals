import pandas as pd
from datetime import datetime

#Import Approvals and Prisma Workbook
approvals = pd.read_excel('Input/Approvals Grid.xlsx', sheet_name = 0)
prisma = pd.read_excel('Input/Prisma.xlsx', sheet_name = 0)

prisma = prisma.rename(columns={'Month of service': 'Month of Service', 'Estimate code':'Est', 'Supplier name': 'Publisher','Billed':'Ordered'})

#Convert Month of Service to datetime
approvals['Month of Service'] = pd.to_datetime(approvals['Month of Service'])
prisma['Month of Service'] = pd.to_datetime(prisma['Month of Service'])

# #Concatenate Estimate, MoS, Supplier
# approvals['UI'] = approvals['Est'].map(str) + approvals['Month of Service'].map(str) + approvals['Publisher']
# prisma['UI'] = prisma['Est'].map(str) + prisma['Month of Service'].map(str) + prisma['Publisher']

match = approvals.merge(prisma, how='inner',left_on=['Est','Month of Service','Publisher'],right_on=['Est','Month of Service','Publisher'],indicator=True)

match['Approval Status'] = ''

for i, j in match.iterrows():
    if(match['Product Name'][i] != 'Fees'):
        if(match['Ordered'][i] == match['Actual Net Billable'][i]):
            match.at[i,'Approval Status'] = 'Media Buy: Approved'
        else:
            match.at[i,'Approval Status'] = 'Media Buy: Ordered <> Billable'
    else:
        match.at[i, 'Approval Status'] = 'Fee Buy: Unknown'
match.to_csv('test.csv')