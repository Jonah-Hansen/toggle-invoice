to get started, please run
`pip install -r requirements.txt -t ./venv`

rename `preference.example.json` to `preference.json`
and populate with your info.

running 
`python main.py`
will generate a pdf inside /generated_pdfs for the most recent pay day as defined in preferences > options > payDays
pdf will include the hours from projects in your workspace, rounded to preferences > options > roundingMinutes