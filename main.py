import tkinter as tk
from tkinter import filedialog, messagebox
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
        due = due.replace(day=1, month=1 if due.month == 12 else due.month + 1,
                          year=due.year + 1 if due.month == 12 else due.year) - timedelta(days=1)
    if terms == "21MFI":
        due = due.replace(day=21, month=1 if due.month == 12 else due.month +
                          1, year=due.year + 1 if due.month == 12 else due.year)
    return due.strftime('%d/%m/%Y')


def generate_invoice():
    # Retrieve data from GUI fields
    me = {
        'companyName': companyName_entry.get(),
        'name': name_entry.get(),
        'address': {
            'streetAddress': streetAddress_entry.get(),
            'city': city_entry.get(),
            'province': province_entry.get(),
            'country': country_entry.get(),
            'postalCode': postalCode_entry.get()
        },
        'email': email_entry.get(),
        'phone': phone_entry.get(),
        'taxID': taxID_entry.get(),
        'apiKey': apiKey_entry.get(),
        'workspaceId': workspaceId_entry.get()
    }

    client = {
        'name': clientName_entry.get(),
        'companyName': clientCompanyName_entry.get(),
        'address': {
            'streetAddress': clientStreetAddress_entry.get(),
            'city': clientCity_entry.get(),
            'province': clientProvince_entry.get(),
            'country': clientCountry_entry.get(),
            'postalCode': clientPostalCode_entry.get()
        }
    }

    options = {
        'invoiceNumber': int(invoiceNumber_entry.get()),
        'paymentTerms': paymentTerms_entry.get(),
        'roundingMinutes': int(roundingMinutes_entry.get()),
        'rate': float(rate_entry.get()),
        'payDays': list(map(int, payDays_entry.get().split(','))),
        'taxType': str(taxType_entry.get()),
        'taxPercent': int(taxPercent_entry.get())
    }

    today = date.today()
    invoiceNumber = options['invoiceNumber']
    name = me['companyName'] if len(me['companyName']) > 0 else me['name']

    # Ask user where to save the file
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                             filetypes=[
                                                 ("PDF files", "*.pdf")],
                                             initialfile=f"{name.replace(' ', '')}Invoice{today.strftime('%Y-%m-%d')}_{invoiceNumber}.pdf")

    if not file_path:
        return  # User cancelled save dialog

    c = Canvas(file_path, pagesize=LETTER, bottomup=False)

    # Title
    c.setFont("Helvetica-Bold", 20)
    c.drawString(inch, inch, name)
    c.drawRightString(c._pagesize[0] - inch, inch, 'INVOICE')

    # My details
    sectionY = 1.25
    addr = me['address']
    c.setFont('Helvetica', 11)
    if len(addr['streetAddress']) > 0:
        sectionY += .25
        c.drawString(inch, sectionY * inch, addr['streetAddress'])
    if len(addr['city'] + addr['province'] + addr['country']) > 0:
        sectionY += .25
        c.drawString(inch, (sectionY) * inch,
                     f"{addr['city']} {addr['province']}, {addr['country']} {addr['postalCode']}")
    if len(me['email']) > 0:
        sectionY += .25
        c.drawString(inch, sectionY * inch, me['email'])
    if len(me['phone']) > 0:
        sectionY += .25
        c.drawString(inch, sectionY * inch, me['phone'])
    if len(me['taxID']) > 0:
        sectionY += .25
        c.drawString(inch, sectionY * inch, f"BN: {me['taxID']}")

    # Client details
    sectionY = 3
    c.setFont('Helvetica-Bold', 11)
    c.drawString(inch, sectionY * inch, 'Bill To:')
    addr = client['address']
    c.setFont('Helvetica', 11)
    if len(client['name']) > 0:
        sectionY += .25
        c.drawString(inch, sectionY * inch, client['name'])
    if len(client['companyName']) > 0:
        sectionY += .25
        c.drawString(inch, (sectionY) * inch, client['companyName'])
    if len(addr['streetAddress']) > 0:
        sectionY += .25
        c.drawString(inch, sectionY * inch, addr['streetAddress'])
    if len(addr['city'] + addr['province'] + addr['country']) > 0:
        sectionY += .25
        c.drawString(inch, sectionY * inch,
                     f"{addr['city']} {addr['province']}, {addr['country']} {addr['postalCode']}")

    # Invoice details
    c.setFont('Helvetica-Bold', 11)
    c.drawRightString(c._pagesize[0] - 2 * inch, 3.20 * inch, 'Invoice #')
    c.drawRightString(c._pagesize[0] - 2 * inch, 3.50 * inch, 'Invoice Date')
    c.drawRightString(c._pagesize[0] - 2 * inch, 3.80 * inch, 'Due Date')

    c.setFont('Helvetica', 11)
    c.drawRightString(c._pagesize[0] - 1 * inch,
                      3.20 * inch, str(options['invoiceNumber']))
    c.drawRightString(c._pagesize[0] - 1 * inch,
                      3.50 * inch, today.strftime('%d/%m/%Y'))
    c.drawRightString(c._pagesize[0] - 1 * inch, 3.80 *
                      inch, getDueDate(today, options['paymentTerms']))

    # Get data
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

    subtotal = 0
    for work in times:
        project = requests.get(
            f"https://api.track.toggl.com/api/v9/workspaces/{me['workspaceId']}/projects/{work['project_id']}",
            headers={'content-type': 'application/json',
                     'Authorization': f'Basic {auth}'}
        ).json()
        hours = round(options['roundingMinutes'] * round(
            (sum(work['seconds']) / 60) / options['roundingMinutes']) / 60, 2)
        total = round(hours * options['rate'], 2)
        subtotal += total
        data.append([
            project['name'],
            hours,
            '{0:.2f}'.format(options['rate']),
            '{0:.2f}'.format(total)
        ])

    data.append(['', '', 'SUBTOTAL', '{0:.2f}'.format(
        subtotal)])
    taxes = subtotal * (options["taxPercent"]/100)
    data.append(['', '', options["taxType"], '{0:.2f}'.format(
        taxes)])
    data.append(['', '', 'TOTAL', '{0:.2f}'.format(
        subtotal + taxes)])

    # Items
    c.saveState()
    c.bottomup = True
    c.transform(1, 0, 0, -1, 0, c._pagesize[1])
    # Configure style and word wrap
    items = Table(data, colWidths=[
                  3.75 * inch, inch, .75 * inch, inch], rowHeights=.35 * inch, vAlign='MIDDLE')
    style = TableStyle(
        [('GRID', (0, 0), (-1, -4), .75, '#505050'),
         ('GRID', (-1, -4), (-1, -1), .75, '#505050'),
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
    items.drawOn(c, inch, c._pagesize[1] - 6 * inch)

    c.restoreState()
    c.save()

    # Increment invoice number in preferences
    options['invoiceNumber'] += 1
    prefs['options']['invoiceNumber'] = options['invoiceNumber']
    with open('./preferences.json', 'w') as f:
        json.dump(prefs, f)
        f.close()

    messagebox.showinfo("Success", "Invoice generated successfully!")
    exit()


# Load preferences
with open('./preferences.json') as f:
    prefs = json.load(f)

# GUI setup
root = tk.Tk()
root.title("Invoice Generator")

# My Info
tk.Label(root, text="My Info", font=('Helvetica', 14, 'bold')).grid(
    row=0, columnspan=2, pady=10)
tk.Label(root, text="Company Name").grid(row=1, column=0, sticky=tk.E)
companyName_entry = tk.Entry(root)
companyName_entry.grid(row=1, column=1, padx=10, pady=5)
companyName_entry.insert(0, prefs['myInfo'].get('companyName', ''))

tk.Label(root, text="Name").grid(row=2, column=0, sticky=tk.E)
name_entry = tk.Entry(root)
name_entry.grid(row=2, column=1, padx=10, pady=5)
name_entry.insert(0, prefs['myInfo'].get('name', ''))

tk.Label(root, text="Street Address").grid(row=3, column=0, sticky=tk.E)
streetAddress_entry = tk.Entry(root)
streetAddress_entry.grid(row=3, column=1, padx=10, pady=5)
streetAddress_entry.insert(
    0, prefs['myInfo']['address'].get('streetAddress', ''))

tk.Label(root, text="City").grid(row=4, column=0, sticky=tk.E)
city_entry = tk.Entry(root)
city_entry.grid(row=4, column=1, padx=10, pady=5)
city_entry.insert(0, prefs['myInfo']['address'].get('city', ''))

tk.Label(root, text="Province").grid(row=5, column=0, sticky=tk.E)
province_entry = tk.Entry(root)
province_entry.grid(row=5, column=1, padx=10, pady=5)
province_entry.insert(0, prefs['myInfo']['address'].get('province', ''))

tk.Label(root, text="Country").grid(row=6, column=0, sticky=tk.E)
country_entry = tk.Entry(root)
country_entry.grid(row=6, column=1, padx=10, pady=5)
country_entry.insert(0, prefs['myInfo']['address'].get('country', ''))

tk.Label(root, text="Postal Code").grid(row=7, column=0, sticky=tk.E)
postalCode_entry = tk.Entry(root)
postalCode_entry.grid(row=7, column=1, padx=10, pady=5)
postalCode_entry.insert(0, prefs['myInfo']['address'].get('postalCode', ''))

tk.Label(root, text="Email").grid(row=8, column=0, sticky=tk.E)
email_entry = tk.Entry(root)
email_entry.grid(row=8, column=1, padx=10, pady=5)
email_entry.insert(0, prefs['myInfo'].get('email', ''))

tk.Label(root, text="Phone").grid(row=9, column=0, sticky=tk.E)
phone_entry = tk.Entry(root)
phone_entry.grid(row=9, column=1, padx=10, pady=5)
phone_entry.insert(0, prefs['myInfo'].get('phone', ''))

tk.Label(root, text="tax ID").grid(row=10, column=0, sticky=tk.E)
taxID_entry = tk.Entry(root)
taxID_entry.grid(row=10, column=1, padx=10, pady=5)
taxID_entry.insert(0, prefs['myInfo'].get('taxID', ''))

tk.Label(root, text="API Key").grid(row=11, column=0, sticky=tk.E)
apiKey_entry = tk.Entry(root)
apiKey_entry.grid(row=11, column=1, padx=10, pady=5)
apiKey_entry.insert(0, prefs['myInfo'].get('apiKey', ''))

tk.Label(root, text="Workspace ID").grid(row=12, column=0, sticky=tk.E)
workspaceId_entry = tk.Entry(root)
workspaceId_entry.grid(row=12, column=1, padx=10, pady=5)
workspaceId_entry.insert(0, prefs['myInfo'].get('workspaceId', ''))

# Client Info
tk.Label(root, text="Client Info", font=('Helvetica', 14, 'bold')).grid(
    row=13, columnspan=2, pady=10)
tk.Label(root, text="Client Name").grid(row=14, column=0, sticky=tk.E)
clientName_entry = tk.Entry(root)
clientName_entry.grid(row=14, column=1, padx=10, pady=5)
clientName_entry.insert(0, prefs['clientInfo'].get('name', ''))

tk.Label(root, text="Client Company Name").grid(row=15, column=0, sticky=tk.E)
clientCompanyName_entry = tk.Entry(root)
clientCompanyName_entry.grid(row=15, column=1, padx=10, pady=5)
clientCompanyName_entry.insert(0, prefs['clientInfo'].get('companyName', ''))

tk.Label(root, text="Client Street Address").grid(
    row=16, column=0, sticky=tk.E)
clientStreetAddress_entry = tk.Entry(root)
clientStreetAddress_entry.grid(row=16, column=1, padx=10, pady=5)
clientStreetAddress_entry.insert(
    0, prefs['clientInfo']['address'].get('streetAddress', ''))

tk.Label(root, text="Client City").grid(row=17, column=0, sticky=tk.E)
clientCity_entry = tk.Entry(root)
clientCity_entry.grid(row=17, column=1, padx=10, pady=5)
clientCity_entry.insert(0, prefs['clientInfo']['address'].get('city', ''))

tk.Label(root, text="Client Province").grid(row=18, column=0, sticky=tk.E)
clientProvince_entry = tk.Entry(root)
clientProvince_entry.grid(row=18, column=1, padx=10, pady=5)
clientProvince_entry.insert(
    0, prefs['clientInfo']['address'].get('province', ''))

tk.Label(root, text="Client Country").grid(row=19, column=0, sticky=tk.E)
clientCountry_entry = tk.Entry(root)
clientCountry_entry.grid(row=19, column=1, padx=10, pady=5)
clientCountry_entry.insert(
    0, prefs['clientInfo']['address'].get('country', ''))

tk.Label(root, text="Client Postal Code").grid(row=20, column=0, sticky=tk.E)
clientPostalCode_entry = tk.Entry(root)
clientPostalCode_entry.grid(row=20, column=1, padx=10, pady=5)
clientPostalCode_entry.insert(
    0, prefs['clientInfo']['address'].get('postalCode', ''))

# Invoice Options
tk.Label(root, text="Invoice Options", font=(
    'Helvetica', 14, 'bold')).grid(row=21, columnspan=2, pady=10)
tk.Label(root, text="Invoice Number").grid(row=22, column=0, sticky=tk.E)
invoiceNumber_entry = tk.Entry(root)
invoiceNumber_entry.grid(row=22, column=1, padx=10, pady=5)
invoiceNumber_entry.insert(0, prefs['options'].get('invoiceNumber', ''))

tk.Label(root, text="Payment Terms").grid(row=23, column=0, sticky=tk.E)
paymentTerms_entry = tk.Entry(root)
paymentTerms_entry.grid(row=23, column=1, padx=10, pady=5)
paymentTerms_entry.insert(0, prefs['options'].get('paymentTerms', ''))

tk.Label(root, text="Rounding Minutes").grid(row=24, column=0, sticky=tk.E)
roundingMinutes_entry = tk.Entry(root)
roundingMinutes_entry.grid(row=24, column=1, padx=10, pady=5)
roundingMinutes_entry.insert(0, prefs['options'].get('roundingMinutes', ''))

tk.Label(root, text="Rate").grid(row=25, column=0, sticky=tk.E)
rate_entry = tk.Entry(root)
rate_entry.grid(row=25, column=1, padx=10, pady=5)
rate_entry.insert(0, prefs['options'].get('rate', ''))

tk.Label(root, text="Pay Days").grid(row=26, column=0, sticky=tk.E)
payDays_entry = tk.Entry(root)
payDays_entry.grid(row=26, column=1, padx=10, pady=5)
payDays_entry.insert(0, ','.join(
    map(str, prefs['options'].get('payDays', []))))

tk.Label(root, text="taxes", font=(
    'Helvetica', 12, 'bold')).grid(row=27, columnspan=1, pady=2)
tk.Label(root, text="Type").grid(row=28, column=0, sticky=tk.E)
taxType_entry = tk.Entry(root)
taxType_entry.grid(row=28, column=1, padx=10, pady=5)
taxType_entry.insert(0, prefs['options'].get('taxes').get("type", ""))
tk.Label(root, text="percent").grid(row=29, column=0, sticky=tk.E)
taxPercent_entry = tk.Entry(root)
taxPercent_entry.grid(row=29, column=1, padx=10, pady=5)
taxPercent_entry.insert(0, prefs['options'].get('taxes').get("percent", 0))

# Generate Invoice Button
generate_button = tk.Button(
    root, text="Generate Invoice", command=generate_invoice)
generate_button.grid(row=30, columnspan=2, pady=20)

root.mainloop()
