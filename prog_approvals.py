import pandas as pd
from datetime import date
import time
from progress_bar import printProgressBar

now = date.today().strftime('%B_%d_%Y')

prisma_cols = ['Month of service', 'Estimate code', 'Supplier short name', 'Actual Net Billable', 'Placement ID', 'Placement name']

#Import Approvals and Prisma Workbook
approvals = pd.read_excel('Input/Approvals Grid.xlsx', sheet_name = 0)
prisma = pd.read_excel('Input/Prisma.xlsx', header=1, sheet_name = 0, usecols=prisma_cols)
# fees_sums = pd.read_excel('summed sheet')

prisma = prisma.rename(columns={'Month of service': 'Month of Service', 'Estimate code':'Est', 'Supplier short name': 'Publisher'})

#Convert Month of Service to datetime
approvals['Month of Service'] = pd.to_datetime(approvals['Month of Service'])
prisma['Month of Service'] = pd.to_datetime(prisma['Month of Service'])

match = approvals.merge(prisma, how='left',on=['Est','Month of Service','Publisher'])

match['Approval Status'] = ''
match['Estimate Status'] = ''

print('Rows found in Draft Approvals Sheet: ' + str(approvals['Est'].count()))
print('Rows found in Prisma Sheet: ' + str(prisma['Est'].count()))
print('Rows found in Match Sheet: ' + str(match['Est'].count()))

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

items = list(range(0,match['Est'].count()))
l = len(items)

# printProgressBar(0, l, prefix = 'Progress:', suffix = '', length = 50)

for i, j in match.iterrows():
    # for t, item in enumerate(items):
    if('Media Buy (Ordered <> Billable)' in match['Approval Status'][i]):
        estimate = match['Est'][i]
        for k, w in match.iterrows():
            if(match['Est'][k] == estimate):
                match.at[k,'Estimate Status'] = 'Dependent'
    print(str(i + 1) + '/' + str(match['Est'].count()) + ' Complete')
        # time.sleep(0.1)
        # printProgressBar(t + 1, l, prefix = 'Progress:', suffix = str(t + 1) + '/' + str(match['Est'].count()), length = 50)
       
# fees_sums

# Estimate Key against a summed sheet
# 	If fee is reconciled or fee does not have a previous month's billing, check Ordered vs Billed
# 	If fee isn't reconciled, then fee must be within $300 +/-, previous month OR a delta (Calculated Fee - previous month's ordered)


draft_approvals_output = 'draft_approvals_output_' + now + '.csv'

match.to_csv('Output/' + draft_approvals_output)