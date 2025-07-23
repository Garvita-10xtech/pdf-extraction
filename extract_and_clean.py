import fitz  # PyMuPDF
import pandas as pd
import re

# === Step 1: Extract from PDF ===
pdf_path = "Aahar 2025 Fair Guide.pdf"
doc = fitz.open(pdf_path)

entries = []
for page in doc:
    blocks = page.get_text("blocks")
    blocks = sorted(blocks, key=lambda b: (b[1], b[0]))

    for b in blocks:
        text = b[4].strip()
        if not text:
            continue
        if "@" in text and len(text.splitlines()) >= 3:
            entries.append(text)

data = []
for entry in entries:
    lines = entry.splitlines()

    # Smart Company Name detection
    company_name = ""
    for line in lines:
        if not line.strip():
            continue
        if any(kw in line.lower() for kw in ["phone", "contact", "email", "regional office", "head office", "tel", "fax"]):
            continue
        if re.match(r'^\d+$', line.strip()):
            continue
        company_name = line.strip()
        break

    email = ''
    phone = ''
    contact = ''
    address_lines = []
    products = []

    for line in lines:
        lower = line.lower()

        if '@' in line and not email:
            match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', line)
            if match:
                email = match.group(0)

        elif any(w in lower for w in ['tel', 'mobile', 'phone']) and not phone:
            match = re.search(r'\+?\d[\d\s\-().]{6,}', line)
            if match:
                phone = match.group(0)

        elif 'contact person' in lower and not contact:
            contact = line.split(":")[-1].strip()

        elif any(w in lower for w in ['product', 'line', 'equipment']):
            products.append(line.strip())

        elif not any(k in lower for k in ['email', 'tel', 'contact', 'product']):
            address_lines.append(line.strip())

    data.append({
        "Company Name": company_name,
        "Address": " ".join(address_lines),
        "Contact Person": contact,
        "Tel./Mobile": phone,
        "E-mail": email,
        "Products on Display": ", ".join(products)
    })

# === Step 2: Create Excel File ===
df = pd.DataFrame(data)
intermediate_path = "aahar_companies_raw.xlsx"
df.to_excel(intermediate_path, index=False)

# === Step 3: Clean the Excel ===

# 1. Remove Company Name rows that look like phone numbers
phone_like_pattern = re.compile(r'^\+?\d[\d\s\-().]{6,}$')
df = df[~df["Company Name"].astype(str).str.match(phone_like_pattern)]

# 2. Remove duplicates by company name
df = df.drop_duplicates(subset=["Company Name"], keep="first")

# 3. Remove invalid email entries
# === Step 3: Clean the Excel ===

# Pattern to identify phone numbers and emails
phone_like_pattern = re.compile(r'^\+?\d[\d\s\-().]{6,}$')
email_like_pattern = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')
phone_anywhere_pattern = re.compile(r'\+?\d[\d\s\-().]{6,}')

# 1. Remove rows where Company Name is a phone number or email or contains one
df = df[~df["Company Name"].astype(str).str.match(phone_like_pattern)]
df = df[~df["Company Name"].astype(str).str.contains(email_like_pattern)]
df = df[~df["Company Name"].astype(str).str.contains(phone_anywhere_pattern)]

# 2. Remove duplicates by company name
df = df.drop_duplicates(subset=["Company Name"], keep="first")

# 3. Remove invalid email entries
df = df[df["E-mail"].astype(str).str.contains(r'^[\w\.-]+@[\w\.-]+\.\w+$', na=False)]

# 4. Normalize phone numbers
df["Tel./Mobile"] = df["Tel./Mobile"].astype(str).str.replace(r'[\s\-\(\)]', '', regex=True)

# 5. Split products
df["Product List"] = df["Products on Display"].fillna("").apply(
    lambda x: [item.strip() for item in x.split(",") if item.strip()]
)

# === Step 4: Save Final Cleaned Excel ===
final_path = "aahar_companies_cleaned_final.xlsx"
df.to_excel(final_path, index=False)

print(f"✅ All done! Cleaned data saved to: {final_path}")
