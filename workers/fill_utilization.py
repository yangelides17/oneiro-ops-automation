"""
fill_utilization.py — Monthly Workforce Utilization Form filler (dummy data test)
"""
import logging, warnings
logging.getLogger('pypdf').setLevel(logging.ERROR)
warnings.filterwarnings('ignore', message='.*Font dictionary.*not found.*', module='pypdf')

from pypdf import PdfReader, PdfWriter

# Painter row — Journey-Level, Helper, Apprentice, Trainee, New Hires, Layoffs
# Each section: TOT B H A NA FEM
PAINTER_FIELDS = {
    'jl':  ['painter_jl_tot','painter_jl_b','painter_jl_h','painter_jl_a','painter_jl_na','painter_jl_fem'],
    'hlp': ['page1_field126','page1_field96','page1_field3','page1_field50','page1_field4','page1_field5'],
    'app': ['page1_field132','page1_field53','page1_field54','page1_field55','page1_field56','page1_field6'],
    'tra': ['page1_field7','page1_field8','page1_field9','page1_field10','page1_field11','page1_field12'],
    'new': ['page1_field13','page1_field14','page1_field15','page1_field16','page1_field17','page1_field18'],
    'lay': ['page1_field19','page1_field20','page1_field21','page1_field22','page1_field23','page1_field24'],
}

DUMMY = {
    "month":            "August",
    "year":             "2025",
    "dls_file_number":  "",
    "percent_complete": "",
    "contractor":       "Oneiro Collection LLC",
    "contract_number":  "84122MBTP496",
    "comments":         "",
    "contractor_name":  "Marianthy Angelides",
    "contractor_title": "President",
    "date":             "10/13/2025",
    # Painter row — [TOT, B, H, A, NA, FEM] per section
    "painter": {
        "jl":  ["2", "1", "", "", "", ""],
        "hlp": ["",  "",  "", "", "", ""],
        "app": ["",  "",  "", "", "", ""],
        "tra": ["",  "",  "", "", "", ""],
        "new": ["",  "",  "", "", "", ""],
        "lay": ["",  "",  "", "", "", ""],
    }
}

def fill(data: dict, template_path: str, output_path: str) -> str:
    reader = PdfReader(template_path)
    writer = PdfWriter()
    writer.append(reader)

    f = {}
    f['Month']            = data['month']
    f['Year']             = data['year']
    f['DLS_file_number']  = data.get('dls_file_number', '')
    f['Percent_Complete'] = data.get('percent_complete', '')
    f['Contractor']       = data['contractor']
    f['Contract Number']  = data['contract_number']
    f['comments']         = data.get('comments', '')
    f['contractor_name']  = data['contractor_name']
    f['contractor_title'] = data['contractor_title']
    f['date']             = data['date']

    painter = data.get('painter', {})
    for section, field_names in PAINTER_FIELDS.items():
        values = painter.get(section, ['','','','','',''])
        for field_name, val in zip(field_names, values):
            f[field_name] = str(val)

    # Fill across all pages
    for page in writer.pages:
        writer.update_page_form_field_values(page, f)

    with open(output_path, 'wb') as fh:
        writer.write(fh)
    print(f'✅  Filled → {output_path}')
    return output_path

if __name__ == '__main__':
    fill(
        DUMMY,
        'mnt/uploads/Monthly Employee Utilization Form Template_FORM.pdf',
        'mnt/oneiro-ops-script/_tmp_fills/Utilization_TEST.pdf'
    )
