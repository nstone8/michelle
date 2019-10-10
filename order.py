import pickle
import os.path
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import openpyxl
import time
import datetime

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1pCLpRQlAclu3I6od-C1gUeZsOhkaOwHEOy-bzrnY80A'
#SAMPLE_RANGE_NAME = 'Class Data!A2:E'
creds=None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range='Expenses').execute()
values = result.get('values', [])

date=[]
project=[]
supplier=[]
description=[]
business_purpose=[]
product_no=[]
qty=[]
unit_cost=[]
total_cost=[]
unit_count=[]
unit_size=[]
need_date=[]
ordered=[]
received=[]
ordered_for=[]
notes=[]
for v in values:
    date.append(v[0])
    project.append(v[1])
    supplier.append(v[2])
    description.append(v[3])
    business_purpose.append(v[4])
    product_no.append(v[5])
    qty.append(v[6])
    unit_cost.append(v[7])
    total_cost.append(v[8])
    unit_count.append(v[9])
    unit_size.append(v[10])
    need_date.append(v[11])
    ordered.append(v[12])
    received.append(v[13])

    new_total_cost=[]
    for t in total_cost:
        if t:
            if t[0]=='$':
                t=t[1:]
        else:
            t=None
        try:
            t=float(t)
        except ValueError:
            t=None
        except TypeError:
            t=None
        new_total_cost.append(t)
        
    try:
        ordered_for.append(v[13])
    except IndexError:
        ordered_for.append(None)
    try:
        notes.append(v[14])
    except IndexError:
        notes.append(None)
total_cost=new_total_cost
frame=pd.DataFrame(dict(
    date=date,
    project=project,
    supplier=supplier,
    description=description,
    business_purpose=business_purpose,
    product_no=product_no,
    qty=qty,
    unit_cost=unit_cost,
    total_cost=total_cost,
    unit_count=unit_count,
    unit_size=unit_size,
    need_date=need_date,
    ordered=ordered,
    received=received,
    ordered_for=ordered_for,
    notes=notes))

not_ordered=frame.loc[frame.loc[:,'ordered']=='',:]

suppliers=set(not_ordered.loc[:,'supplier'])


for s in suppliers:
    this_supplier=not_ordered.loc[not_ordered.loc[:,'supplier']==s,:]
    #Load Template
    wb=openpyxl.load_workbook('Purchase Request Form.xlsx')
    sheet=wb.active
    date_str='{1}/{2}/{0}'.format(*time.localtime()[0:3])
    #fill in date
    sheet['K8'].value=date_str
    #fill in vendor
    sheet['B14'].value=s
    #fill in date needed
    item_no_cells=['C'+str(r) for r in range(22,37)]
    descr_cells=['D'+str(r) for r in range(22,37)]
    quant_cells=['N'+str(r) for r in range(22,37)]
    unit_price_cells=['P'+str(r) for r in range(22,37)]
    ext_cells=['S'+str(r) for r in range(22,37)]
    account_cells=['B'+str(r) for r in range(53,58)]
    cost_cells=['D'+str(r) for r in range(53,58)]
    
    item_numbers=this_supplier.loc[:,'product_no']
    descriptions=this_supplier.loc[:,'description']
    quants=this_supplier.loc[:,'qty']
    unit_prices=this_supplier.loc[:,'unit_cost']
    extensions=this_supplier.loc[:,'total_cost']
    bp=this_supplier.loc[:,'business_purpose']
    account_cost=this_supplier.loc[:,['project','total_cost']].groupby('project').agg(sum).reset_index()
    accounts=account_cost.loc[:,'project']
    costs=account_cost.loc[:,'total_cost']
    date_needed=this_supplier.loc[:,'need_date']
    these_dates=[]
    for date in date_needed:
    	if not date: continue
    	this_month,this_day,this_year=date.split('/')
    	this_date=datetime.date(month=int(this_month),day=int(this_day),year=2000+int(this_year))
    	these_dates.append(time.mktime(this_date.timetuple()))
    if these_dates:
    	date_needed_str=time.strftime('%m/%d/%Y',time.localtime(min(these_dates)))
    else:
    	date_needed_sec=time.time()+14*24*60*60
    	date_needed_str=time.strftime('%m/%d/%Y',time.localtime(date_needed_sec))
    sheet['N37'].value=date_needed_str
    if len(item_numbers)>len(item_no_cells):
        raise Exception('Too many items to fit on one form for supplier {}'.format(s))

    for icell,inum in zip(item_no_cells,item_numbers):
        sheet[icell].value=inum

    for desc_cell,desc in zip(descr_cells,descriptions):
        sheet[desc_cell].value=desc

    for quant_cell,quant in zip(quant_cells,quants):
        sheet[quant_cell].value=quant

    for uprice_cell,uprice in zip(unit_price_cells,unit_prices):
        if uprice[0]=='$':
            uprice=uprice[1:]
            
        sheet[uprice_cell].value=float(uprice)

    for ext_cell,ext in zip(ext_cells,extensions):
        sheet[ext_cell].value=ext

    for account_cell,account in zip(account_cells,accounts):
        sheet[account_cell].value=account

    for cost_cell,cost in zip(cost_cells,costs):
        sheet[cost_cell].value=cost
        
    #fill in business purpose
    sheet['E47'].value=', '.join(bp)
    wb.save(os.path.join('Order_sheets','{} order.xlsx'.format(s)))
        
