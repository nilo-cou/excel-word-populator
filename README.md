# 🎨 excel-word-populator

> 📋 Focuses on the mechanism—taking data from Excel and "populating" or "filling" the Word template.

---

## 🚀 Quick Overview

This Python script automates the generation of Word documents by:
- 📊 Reading data from Excel spreadsheets
- 🔄 Mapping Excel columns to Word template placeholders
- 📄 Generating personalized Word documents for each row
- 🎯 Organizing output with timestamped directories

---

## 📦 Dependencies

```python
import pandas as pd           # 📊 Data manipulation
from docxtpl import DocxTemplate  # 📝 Word template rendering
from docx.shared import Pt    # 🎨 Font formatting
from numbers import Number    # 🔢 Numeric type checking
import os                     # 📁 File system operations
import re                     # 🔍 Pattern matching
from datetime import datetime # ⏰ Timestamping
```

---

## ⚙️ Configuration

### 🔧 File Settings

Update these paths to match your actual files:

```python
excel_file = r'C:\Users\ndeka\OneDrive\Desktop\Nikitas pickle excel\1 Northeast Crafthouse for bios Generation.xlsx'
template_file = r'C:\Users\ndeka\OneDrive\Desktop\Nikitas pickle excel\Product bios Template.docx'
output_dir = r'C:\Users\ndeka\OneDrive\Desktop\Nikitas pickle excel\Output Directory'
```

✅ **Validation checks:**
- ✓ Excel file exists
- ✓ Word template exists
- ✓ Output directory is created if missing

---

## 🔧 Key Features

### 1️⃣ Smart Column Mapping 🗺️
- 🧹 Cleans Excel headers to match Word placeholders
- 📍 Case-insensitive placeholder matching
- 🎯 Automatically handles common unit tokens (₹, %, Rs., etc.)

### 2️⃣ Dynamic Placeholder Detection 🔎
- 🔍 Extracts all placeholders from your Word template
- 🗺️ Creates best-effort mapping between Excel columns and template variables
- 📊 Handles duplicate normalized placeholders intelligently

### 3️⃣ Numeric Precision 💰
- 💵 Preserves economic values with full precision
- 🔢 Removes unnecessary decimals from integers
- 💯 Detects price, MRP, GST, commission, tax, and discount fields

### 4️⃣ Font Styling ✨
- 🎨 Applies **Montserrat** font at **12pt** to all generated documents
- 📝 Consistent formatting across paragraphs and tables
- ✨ Professional appearance for all outputs

### 5️⃣ Smart Filename Generation 🆔
- 🆔 Uses SKU or Product Code when available
- 🔄 Falls back to row index if identifier missing
- 📌 Normalized filenames (spaces, slashes handled)

### 6️⃣ Output Organization 📂
- ⏰ Creates timestamped run directories: `run_YYYYMMDD_HHMMSS`
- 📋 Generates placeholder reference file
- 📊 Reports total files generated

---

## 🔄 Processing Pipeline

### 📥 Step 1: Data Loading
```
✓ Reads 'Product bios Template' sheet (or first sheet if not found)
✓ Converts NaN values to empty strings
✓ Cleans column headers for Jinja2 compatibility
```

### 🧹 Step 2: Header Cleaning
- Converts special characters to underscores
- Example: `Vendor's Expected Retail Gain (₹) (Ex GST)` → `Vendor_s_Expected_Retail_Gain_Ex_GST`

### 🗺️ Step 3: Placeholder Mapping
- Extracts template variables from Word document
- Normalizes both Excel columns and template placeholders
- Creates intelligent mapping accounting for unit differences

### 📋 Step 4: Context Building
- Builds dictionary from each Excel row
- Maps DataFrame columns to template placeholders
- Enforces default values for critical fields

### 📄 Step 5: Document Generation
- Loads fresh template copy for each row
- Renders template with row context
- Applies font formatting
- Saves with unique filename

---

## 📋 Default Values

If these fields are missing, defaults are provided:

| Field | Default Value |
|-------|---------------|
| 🏷️ **Product_Display_Name** | "Product name not provided" |
| ⚖️ **Packed_Product_Weight_G** | "Weight not provided" |
| 📐 **Product_Dimension_Cm** | "Dimension not provided" |

---

## 📁 Output Structure

```
Output Directory/
└── 📂 run_20231215_143025/
    ├── 📄 available_placeholders.txt
    ├── 📄 SKU_001_1.docx
    ├── 📄 SKU_002_2.docx
    ├── 📄 SKU_003_3.docx
    └── ...
```

---

## 🎯 How It Works

1. **📥 Load Excel data** with smart sheet detection
2. **🔍 Parse Word template** to find all placeholders
3. **🗺️ Create mapping** between Excel columns and template variables
4. **🔁 For each row:**
   - 🔄 Create context dictionary
   - 💰 Format numeric values appropriately
   - 📄 Render template with context
   - 🎨 Apply consistent styling
   - 💾 Save with unique filename

---

## 💡 Smart Features

### 🎯 Category Fallback
If `Category_Level_3` is missing, it's populated from `Sub_Category_2_Ultimate_Category_Functional_Use_Of_The_Goods`

### 🔎 Case-Insensitive Key Lookup
Finds context keys even with different casing or special characters

### 💵 Economic Value Precision
- 💰 Currency fields preserve exact values
- 🔢 Integers display without decimal points
- 📊 Floats maintain full precision

### 🧩 Placeholder Deduplication
Prevents duplicate numeric values in documents by marking secondary placeholders

---

## ✅ Success Indicators

✨ **Console output shows:**
- 📊 Available placeholders for reference
- 📝 Generated filename for each row
- ✅ Success message with output directory
- 🔢 Total count of generated files

📄 **Output file:** `available_placeholders.txt` lists all template variables

---

## 🚨 Error Handling

**Validates:**
- ✅ Excel file exists
- ✅ Word template exists
- ✅ Output directory is writable
- ✅ Data type compatibility
- ✅ Placeholder resolution

---

## 📊 Example Data Flow

```
Excel Row
    ↓
🧹 Clean & Normalize Columns
    ↓
🗺️ Map to Template Placeholders
    ↓
💰 Format Numeric Values
    ↓
📄 Render Word Template
    ↓
🎨 Apply Font Styling
    ↓
💾 Save as SKU_N.docx
```

---

**✨ Generated with ❤️ for batch document creation! ✨**
