import pandas as pd
from datetime import date
import time
import os
import re

now = date.today().strftime('%B_%d_%Y')
input_path = 'Input/'
output_path = os.path.join('Output/','Draft Output_' + now + '/')

if (os.path.isdir(output_path)):
    print(output_path + ' directory already exists')
else:
    os.mkdir(output_path)
    print(output_path  + ' directory created')

prisma_cols = ['Month of service', 'Estimate code', 'Supplier short name', 'Actual Net Billable', 'Placement ID', 'Placement name']

#Import Approvals, Prisma Workbook, Consolidated Fees
approvals = pd.read_excel('Input/Approvals Grid.xlsx', sheet_name = 0)
prisma = pd.read_excel('Input/Prisma.xlsx', header=1, sheet_name = 0, usecols=prisma_cols)
for filename in os.listdir(input_path):
    if re.match('Consolidated Fees*(.+)\.csv', filename):
        consolidated_fees = pd.read_csv(input_path + filename)
        print('Reading Consolidated Fees file: ' + input_path + filename)
    else:
        print('No consolidated fees file found.')
invoices = pd.read_excel('Input/Invoices.xlsx', sheet=0, usecols=['Invoice month','Supplier name','Invoice status','Product code','Estimate code'])

prisma = prisma.rename(columns={'Month of service': 'Month of Service', 'Estimate code':'Est', 'Supplier short name': 'Publisher'})

#Convert Month of Service to datetime
approvals['Month of Service'] = pd.to_datetime(approvals['Month of Service'])
prisma['Month of Service'] = pd.to_datetime(prisma['Month of Service'])
consolidated_fees['Month'] = pd.to_datetime(consolidated_fees['Month'])

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

match.to_csv(output_path + draft_approvals_output)

print('Draft Approvals Output file saved at: ' + output_path + draft_approvals_output)