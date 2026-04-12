"""
fill_signin.py — Employee Daily Sign-In Log filler (dummy data test)
"""
import logging, warnings
logging.getLogger('pypdf').setLevel(logging.ERROR)
warnings.filterwarnings('ignore', message='.*Font dictionary.*not found.*', module='pypdf')

from pypdf import PdfReader, PdfWriter

# Each employee row has 4 generic fields: Time In, Emp Sig, Time Out, Emp Sig
# Mapped by order from the field dump
ROW_FIELDS = [
    ('page0_field9',  'page0_field12', 'page0_field10', 'page0_field13'),  # row 1
    ('page0_field15', 'page0_field16', 'page0_field17', 'page0_field19'),  # row 2
    ('page0_field22', 'page0_field23', 'page0_field24', 'page0_field25'),  # row 3
    ('page0_field28', 'page0_field29', 'page0_field30', 'page0_field31'),  # row 4
    ('page0_field34', 'page0_field35', 'page0_field36', 'page0_field37'),  # row 5
    ('page0_field40', 'page0_field41', 'page0_field42', 'page0_field43'),  # row 6
    ('page0_field46', 'page0_field47', 'page0_field48', 'page0_field49'),  # row 7
    ('page0_field52', 'page0_field53', 'page0_field54', 'page0_field55'),  # row 8
    ('page0_field58', 'page0_field59', 'page0_field60', 'page0_field61'),  # row 9
    ('page0_field64', 'page0_field65', 'page0_field66', 'page0_field67'),  # row 10
    ('page0_field70', 'page0_field72', 'page0_field73', 'page0_field71'),  # row 11
    ('page0_field76', 'page0_field77', 'page0_field78', 'page0_field79'),  # row 12
]

DUMMY = {
    "prime_contractor":    "Metro Express Services",
    "subcontractor":       "Oneiro Collection LLC",
    "contract_number":     "84122MBTP496",
    "address":             "54-35 48th St, Maspeth NY 11378",
    "agency":              "NYCDOT",
    "project_name":        "PT-11930 | Atlantic Ave",
    "date":                "8/16/25",
    "employees": [
        {"name": "Carlos Solorzano", "classification": "SAT", "time_in": "9:15 AM", "time_out": "1:15 PM"},
        {"name": "Stamati Angelides", "classification": "LP",  "time_in": "8:15 AM", "time_out": "1:15 PM"},
    ],
    "contractor_name":     "Stamati Angelides",
    "contractor_title":    "Crew Leader",
    "date_signed":         "8/16/25",
}

def fill(data: dict, template_path: str, output_path: str) -> str:
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    f = {}
    f['Prime_Contractor']       = data['prime_contractor']
    f['Subcontractor']          = data['subcontractor']
    f['Contract_Number']        = data['contract_number']
    f['Contractor_Address']     = data['address']
    f['Agency']                 = data['agency']
    f['Project_Name']           = data['project_name']
    f['Date_Header']            = data['date']
    f['Contractor_Name']        = data['contractor_name']
    f['Contractor_Title']       = data['contractor_title']
    f['Date_Signature_Block']   = data['date_signed']

    for i, emp in enumerate(data['employees'][:12]):
        n = i + 1
        f[f'Employee_Name_{n}']  = emp['name']
        f[f'Employee_Class_{n}'] = emp['classification']
        ti, sig1, to, sig2 = ROW_FIELDS[i]
        f[ti]   = emp.get('time_in', '')
        f[ti]   = emp.get('time_in', '')
        f[to]   = emp.get('time_out', '')
        # Leave signature fields blank — handwritten in real use
        f[sig1] = ''
        f[sig2] = ''

    writer.update_page_form_field_values(writer.pages[0], f)
    with open(output_path, 'wb') as fh:
        writer.write(fh)
    print(f'✅  Filled → {output_path}')
    return output_path

if __name__ == '__main__':
    fill(
        DUMMY,
        'mnt/uploads/Employee Daily Sign In Log Template_FORM.pdf',
        'mnt/oneiro-ops-script/_tmp_fills/SignIn_TEST.pdf'
    )
