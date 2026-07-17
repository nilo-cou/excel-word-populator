# excel-word-populator
#Focuses on the mechanism—taking data from Excel and "populating" or "filling" the Word template.

import pandas as pd
from docxtpl import DocxTemplate
from docx.shared import Pt
from numbers import Number
import os
import re
from datetime import datetime

# File settings - Ensure these match your actual filenames
# Update to your new Excel file and template path
excel_file = r'C:\Users\ndeka\OneDrive\Desktop\Nikitas pickle excel\1 Northeast Crafthouse for bios Generation.xlsx'
template_file = r'C:\Users\ndeka\OneDrive\Desktop\Nikitas pickle excel\Product bios Template.docx'
output_dir = r'C:\Users\ndeka\OneDrive\Desktop\Nikitas pickle excel\Output Directory'

if not os.path.isfile(excel_file):
    raise FileNotFoundError(f"Excel file not found: {excel_file}")
if not os.path.isfile(template_file):
    raise FileNotFoundError(f"Word template not found: {template_file}")
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# create a timestamped run directory so each run is isolated and easy to count
run_dir = os.path.join(output_dir, f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
os.makedirs(run_dir, exist_ok=True)

# 1. Load Data
# Try Product Bios sheet first; if missing, fall back to the first sheet.
try:
    df = pd.read_excel(excel_file, sheet_name='Product bios Template', header=1)
except ValueError:
    df = pd.read_excel(excel_file, header=1)

# Convert NaN values to empty strings (safer for placeholders)
df = df.fillna('')

# 2. Clean headers to match your Word placeholders
def clean_for_jinja(col):
    # Collapse any run of non-alphanumeric characters into a single underscore.
    # E.g. 'Vendor's Expected Retail Gain (₹) (Ex GST)' -> 'Vendor_s_Expected_Retail_Gain_Ex_GST'
    return re.sub(r'\W+', '_', str(col)).strip('_')

df.columns = [clean_for_jinja(c) for c in df.columns]

# Build a mapping from template placeholders to dataframe columns so
# template vars like 'Mrp' or 'Selling_Price_On_Invoice' can match
# column names that include units like '(Rs.)' or '%' etc.
tpl_doc = DocxTemplate(template_file)
tpl_vars = tpl_doc.get_undeclared_template_variables()

def _norm_for_match(s: str) -> str:
    s = re.sub(r"\W+", "", str(s)).lower()
    # strip common unit tokens that differ between Excel headers and template
    s = re.sub(r"(rs|rs\.|₹|%|cm|g|kg|mm|incgst|exgst)", "", s)
    s = s.rstrip('_')
    return s

col_norm_map = { _norm_for_match(c): c for c in df.columns }
tpl_to_col = {}
for v in tpl_vars:
    tpl_to_col[v] = col_norm_map.get(_norm_for_match(v))

print('Template->Column mapping (best-effort):')
for k, v in tpl_to_col.items():
    print('  ', k, '->', v)

# If the template contains multiple placeholders that normalize to the same
# token (e.g. `MRP_Rs` and `Mrp`), prefer the first occurrence and clear
# the others to avoid placing the same numeric value twice in the doc.
norm_tpl_map = {}
for v in tpl_vars:
    n = _norm_for_match(v)
    norm_tpl_map.setdefault(n, []).append(v)

tpl_primary = {}
for n, vars_list in norm_tpl_map.items():
    # keep first, mark others as secondary
    primary = vars_list[0]
    tpl_primary[n] = primary
    for secondary in vars_list[1:]:
        # ensure secondary placeholders will be left empty if present in template
        tpl_to_col[secondary] = None
# If Category Level 3 is not present, populate it from Sub Category 2
if 'Category_Level_3' not in df.columns:
    src = 'Sub_Category_2_Ultimate_Category_Functional_Use_Of_The_Goods'
    if src in df.columns:
        df['Category_Level_3'] = df[src].fillna("")
    else:
        df['Category_Level_3'] = ''

# 2.b Apply font settings to the whole document
# This forces all text in generated docs to use Montserrat, size 12.
# Note: Montserrat must be installed on the system where the docs are opened.

def apply_font_settings(doc, font_name: str = "Montserrat", font_size_pt: float = 12):
    def _set_run_font(run):
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)

    def _set_paragraph_font(paragraph):
        for run in paragraph.runs:
            _set_run_font(run)

    for paragraph in doc.docx.paragraphs:
        _set_paragraph_font(paragraph)

    for table in doc.docx.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _set_paragraph_font(paragraph)


# Helper: find a context key case-insensitively (ignoring non-alphanum)
def find_key_case_insensitive(context, target_key: str):
    def norm(s):
        return re.sub(r'\W+', '', str(s)).lower()
    t = norm(target_key)
    # exact match first
    for k in context.keys():
        if norm(k) == t:
            return k
    # partial/contains match
    for k in context.keys():
        nk = norm(k)
        if nk.startswith(t) or t in nk:
            return k
    return None

def ensure_context_default(context, target_key: str, default: str):
    k = find_key_case_insensitive(context, target_key)
    if k:
        if not str(context.get(k, '')).strip():
            context[k] = default
    else:
        context[target_key] = default


# 3. Process Each Row
print(f"Starting process for {len(df)} rows...")

for index, row in df.iterrows():
    # Load a fresh template copy
    doc = DocxTemplate(template_file)
    
    # Map the Excel row data to the Word placeholders
    context = row.to_dict()

    # Populate any template placeholders from matching dataframe columns
    for tpl_var, col_name in tpl_to_col.items():
        if col_name and tpl_var not in context:
            context[tpl_var] = context.get(col_name, '')

    # Preserve economic numeric precision: convert economy-related keys to strings
    # Keep integers without trailing .0; keep floats with full precision (no rounding)
    econ_pattern = re.compile(r'mrp|price|gst|commission|tax|discount|payout|gain|selling|retail|tds|tcs', re.I)
    for k in list(context.keys()):
        if econ_pattern.search(k):
            v = context.get(k)
            try:
                if v is None or (isinstance(v, str) and v == ''):
                    continue
                if isinstance(v, Number):
                    # If value is effectively an integer, format without decimals
                    fv = float(v)
                    if fv.is_integer():
                        context[k] = str(int(fv))
                    else:
                        context[k] = repr(v)
                else:
                    context[k] = str(v)
            except Exception:
                context[k] = str(v)
    
    if index == 0:
        placeholder_list = sorted(context.keys())
        print("Available placeholders (context keys) for your Word template:")
        for key in placeholder_list:
            print(f"  {{ {key} }}")
        print("\nEnsure your template uses these exact placeholders (case-sensitive).\n")

        # Write placeholders to a text file for your convenience
        with open(os.path.join(run_dir, 'available_placeholders.txt'), 'w', encoding='utf-8') as f:
            for key in placeholder_list:
                f.write(f"{{{{ {key} }}}}\n")

    # Generate unique filename using SKU/ID; fall back to row index if unknown
    sku = None
    # Prefer columns that contain 'sku' or 'product_code' (case-insensitive)
    for col in row.index:
        if pd.notna(row[col]) and str(row[col]).strip():
            lname = str(col).lower()
            if 'sku' in lname or 'product_code' in lname:
                sku = str(row[col]).strip()
                break

    # If still not found, fall back to some common column names if present
    if not sku:
        fallback = ['Product_Code_Sku_Number', 'SKU', 'Product_Code', 'Code', 'ID', 'Product_ID']
        for key in fallback:
            if key in row.index and pd.notna(row[key]) and str(row[key]).strip():
                sku = str(row[key]).strip()
                break

    if not sku:
        sku = f"row_{index + 1}"
    # Normalize SKU: convert slashes to hyphens, spaces to underscore,
    # remove other punctuation, and strip trailing separators.
    sku = re.sub(r'[\\/]+', '-', sku)
    sku = re.sub(r'\s+', '_', sku)
    sku = re.sub(r'[^A-Za-z0-9\-_]+', '', sku)
    sku = sku.strip('_-')

    # Enforce placeholder text for key properties if missing
    ensure_context_default(context, 'Product_Display_Name', 'Product name not provided')
    ensure_context_default(context, 'Packed_Product_Weight_G', 'Weight not provided')
    ensure_context_default(context, 'Product_Dimension_Cm', 'Dimension not provided')

    # Fill the template
    doc.render(context)
    
    apply_font_settings(doc, font_name="Montserrat", font_size_pt=12)
    
    # Normalize SKU by removing any existing trailing numeric suffix (_<digits>),
    # then append the current row index so filenames look like `SKU_<row>.docx`.
    sku_base = re.sub(r'_[0-9]+$', '', sku)
    filename = f"{sku_base}_{index + 1}.docx"
    save_path = os.path.join(run_dir, filename)

    doc.save(save_path)
    print(f"Generated: {filename}")

print(f"\nSuccess! Check the '{run_dir}' folder for your completed files.")

# report how many files were generated in this run
generated_files = [f for f in os.listdir(run_dir) if f.lower().endswith('.docx')]
print(f"Files generated in this run: {len(generated_files)}")
