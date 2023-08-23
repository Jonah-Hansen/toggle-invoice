from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Table, TableStyle
from datetime import date, timedelta, datetime
import json
import requests
from base64 import b64encode


def getDueDate(today, terms):
    due = datetime(today.year, today.month, today.day)
    if terms == "PIA":
        return 'Payment in advance'
    if terms == "NET7":
        due += timedelta(days=7)
    if terms == "NET10":
        due += timedelta(days=10)
    if terms == "NET30":
        due += timedelta(days=30)
    if terms == "NET60":
        due += timedelta(days=60)
    if terms == "NET90":
        due += timedelta(days=90)
    if terms == "EOM":
        due = due.replace(day=1, month=1 if due.month ==
                          12 else due.month + 1, year=due.year + 1 if due.month == 12 else due.year) - timedelta(days=1)
    if terms == "21MFI":
        due = due.replace(day=21, month=1 if due.month ==
                          12 else due.month + 1, year=due.year + 1 if due.month == 12 else due.year)
    return due.strftime('%d/%m/%Y')


f = open('./preferences.json')
prefs = json.load(f)
f.close()

me, client, options = prefs['myInfo'], prefs['clientInfo'], prefs['options']

invoiceNumber = options['invoiceNumber']

name = me['companyName'] if len(
    me['companyName']) > 0 else me['name']

today = date.today()

c = Canvas('./generated_invoices/' +
           f"{name.replace(' ', '')}Invoice{today.strftime('%Y-%m-%d')}_{invoiceNumber}.pdf", pagesize=LETTER, bottomup=False)

# title
c.setFont("Helvetica-Bold", 20)
c.drawString(inch, inch, name)
c.drawRightString(c._pagesize[0] - inch, inch, 'INVOICE')

# my details
sectionY = 1.25
addr = me['address']
c.setFont('Helvetica', 11)
if len(addr['streetAddress']) > 0:
    sectionY += .25
    c.drawString(inch, sectionY*inch, addr['streetAddress'])
if len(addr['city'] + addr['province'] + addr['country']) > 0:
    sectionY += .25
    c.drawString(inch, (sectionY)*inch,
                 f"{addr['city']} {addr['province']}, {addr['country']} {addr['postalCode']}")
if len(me['email']) > 0:
    sectionY += .25
    c.drawString(inch, sectionY*inch, me['email'])
if len(me['phone']) > 0:
    sectionY += .25
    c.drawString(inch, sectionY*inch, me['phone'])

# client details
sectionY = 3
c.setFont('Helvetica-Bold', 11)
c.drawString(inch, sectionY*inch, 'Bill To:')
addr = client['address']
c.setFont('Helvetica', 11)
if len(client['name']) > 0:
    sectionY += .25
    c.drawString(inch, sectionY * inch, client['name'])
if len(client['companyName']) > 0:
    sectionY += .25
    c.drawString(inch, (sectionY)*inch, client['companyName'])
if len(addr['streetAddress']) > 0:
    sectionY += .25
    c.drawString(inch, sectionY*inch, addr['streetAddress'])
if len(addr['city'] + addr['province'] + addr['country']) > 0:
    sectionY += .25
    c.drawString(inch, sectionY*inch,
                 f"{addr['city']} {addr['province']}, {addr['country']} {addr['postalCode']}")

# invoice details
c.setFont('Helvetica-Bold', 11)
c.drawRightString(c._pagesize[0] - 2*inch, 3.20*inch, 'Invoice #')
c.drawRightString(c._pagesize[0] - 2*inch, 3.50*inch, 'Invoice Date')
c.drawRightString(c._pagesize[0] - 2*inch, 3.80*inch, 'Due Date')

c.setFont('Helvetica', 11)
c.drawRightString(c._pagesize[0] - 1*inch, 3.20*inch, str(options['invoiceNumber']))
c.drawRightString(c._pagesize[0] - 1*inch, 3.50 *
                  inch, today.strftime('%d/%m/%Y'))
c.drawRightString(c._pagesize[0] - 1*inch, 3.80 *
                  inch, getDueDate(today, options['paymentTerms']))

# get data
end_date = today
while not (end_date.day in options['payDays']):
    end_date -= timedelta(days=1)
end_date -= timedelta(days=1)
start_date = end_date - timedelta(days=1)
while not (start_date.day in options['payDays']):
    start_date -= timedelta(days=1)

dateRange = {
    'end_date': end_date.strftime('%Y-%m-%d'),
    'start_date': start_date.strftime('%Y-%m-%d'),
}

token = me['apiKey'] + ':api_token'
auth = b64encode(token.encode('ascii')).decode("ascii")
times = requests.post(
    f"https://api.track.toggl.com/reports/api/v3/workspace/{me['workspaceId']}/weekly/time_entries",
    data=json.dumps(dateRange),
    headers={'content-type': 'application/json',
             'Authorization': f'Basic {auth}'}
).json()

data = [["Project", "hours", "rate", "total"]]

for work in times:
    project = requests.get(
        f"https://api.track.toggl.com/api/v9/workspaces/{me['workspaceId']}/projects/{work['project_id']}",
        headers={'content-type': 'application/json',
                 'Authorization': f'Basic {auth}'}
    ).json()
    hours = options['roundingMinutes'] * \
        round((sum(work['seconds'])/60)/options['roundingMinutes'])/60
    data.append([
        project['name'],
        hours,
        options['rate'],
        round(hours * options['rate']+0.005, 2)
    ])

data.append(['', '', 'TOTAL', sum([i[-1] for i in data[1:]])])

# items
c.saveState()
c.bottomup = True
c.transform(1, 0, 0, -1, 0, c._pagesize[1])
# Configure style and word wrap
items = Table(data, colWidths=[3.75*inch,
              inch, .75*inch, inch], rowHeights=.35*inch, vAlign='MIDDLE')
style = TableStyle(
    [('GRID', (0, 0), (-1, -2), .75, '#505050'),
     ('GRID', (-1, -1), (-1, -1), .75, '#505050'),
     ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
     ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
     ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
     ('SIZE', (0, 1), (-1, -2), 12),
     ('FONT', (0, -1), (-1, -1), 'Helvetica-Bold', 14),
     ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 13),
     ('BACKGROUND', (-1, -1), (-1, -1), '#cccccc'),
     ('BACKGROUND', (0, 0), (-1, 0), '#cccccc'),
     ])
items.setStyle(style)
w, h = items.wrapOn(c, 0, 0)
items.drawOn(c, inch, c._pagesize[1] - 6*inch)

c.restoreState()

c.save()


prefs['options']['invoiceNumber'] += 1
with open('./preferences.json', 'w') as f:
    json.dump(prefs,f)
    f.close()
