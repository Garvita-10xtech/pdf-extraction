import fitz  # PyMuPDF
import pandas as pd
import re

# Load PDF
pdf_path = "Aahar 2025 Fair Guide.pdf"
doc = fitz.open(pdf_path)

entries = []

# Go through each page
for page in doc:
    blocks = page.get_text("blocks")  # (x0, y0, x1, y1, "text", block_no, block_type)

    # Sort blocks top-to-bottom, left-to-right
    blocks = sorted(blocks, key=lambda b: (b[1], b[0]))

    for b in blocks:
        text = b[4].strip()
        if not text:
            continue

        # Heuristic: filter valid blocks that seem like company details
        if "@" in text and len(text.splitlines()) >= 3:
            entries.append(text)

# Now parse each entry into structured data
data = []

for entry in entries:
    lines = entry.splitlines()
    company_name = lines[0].strip()

    email = ''
    phone = ''
    contact = ''
    address_lines = []
    products_lines = []

    for line in lines[1:]:
        lower = line.lower()

        if '@' in line:
            match = re.search(r'[\w\.-]+@[\w\.-]+', line)
            if match:
                email = match.group(0)

        elif any(w in lower for w in ['tel', 'mobile', 'phone']):
            match = re.search(r'(\+?\d[\d\s\-().]+)', line)
            if match:
                phone = match.group(1)

        elif 'contact person' in lower:
            contact = line.split(':')[-1].strip()

        elif any(w in lower for w in ['product', 'line', 'equipment']):
            products_lines.append(line.strip())

        else:
            # Assume address unless it's just a label
            if not any(k in lower for k in ['email', 'tel', 'contact', 'product']):
                address_lines.append(line.strip())

    data.append({
        "Company Name": company_name,
        "Address": " ".join(address_lines),
        "Contact Person": contact,
        "Tel./Mobile": phone,
        "E-mail": email,
        "Products on Display": ", ".join(products_lines)
    })

# Export to Excel
df = pd.DataFrame(data)
df.to_excel("aahar_companies_fitz.xlsx", index=False)

print(f"✅ Done! Extracted {len(data)} companies to aahar_companies_fitz.xlsx")
