import fitz  # PyMuPDF
import re
import pandas as pd

# Paths
pdf_path = "Aahar 2025 Fair Guide.pdf"
output_excel = "aahar2025_fixed_companynames.xlsx"

# Open the PDF
doc = fitz.open(pdf_path)

entries = []
lines_all = []

# Read all text from all pages
for page in doc:
    lines = page.get_text().split('\n')
    lines_all.extend([line.strip() for line in lines if line.strip()])

# Go through all lines and look for STALL: to detect start of a company
i = 0
while i < len(lines_all):
    line = lines_all[i]
    
    if line.startswith("STALL:"):
        # Go back to find company name (assume 1 or 2 lines above)
        company_name = ""
        for j in range(i - 1, max(i - 4, -1), -1):
            if "HALL:" not in lines_all[j] and "STALL:" not in lines_all[j]:
                company_name = lines_all[j]
                break

        # Initialize fields
        address = ""
        contact_person = ""
        phone = ""
        email = ""
        products = ""

        # Look forward for other fields
        j = i + 1
        while j < len(lines_all) and not lines_all[j].startswith("STALL:"):
            if lines_all[j].startswith("Address:"):
                addr_lines = [lines_all[j].replace("Address:", "").strip()]
                j += 1
                while j < len(lines_all) and not any(lines_all[j].startswith(p) for p in ["Contact Person", "Tel", "E-mail", "Products on Display", "STALL:"]):
                    addr_lines.append(lines_all[j])
                    j += 1
                address = " ".join(addr_lines)
                continue

            elif lines_all[j].startswith("Contact Person"):
                contact_person = lines_all[j].split(":", 1)[-1].strip()

            elif lines_all[j].startswith("Tel") or "Mobile" in lines_all[j]:
                phone = lines_all[j].split(":", 1)[-1].strip()

            elif lines_all[j].startswith("E-mail"):
                candidate = lines_all[j].split(":", 1)[-1].strip()
                if re.match(r"[\w\.-]+@[\w\.-]+\.\w+", candidate):
                    email = candidate

            elif lines_all[j].startswith("Products on Display"):
                prod_lines = [lines_all[j].split(":", 1)[-1].strip()]
                j += 1
                while j < len(lines_all) and not any(lines_all[j].startswith(p) for p in ["Contact Person", "Address", "Tel", "E-mail", "STALL:"]):
                    prod_lines.append(lines_all[j])
                    j += 1
                products = " ".join(prod_lines)
                continue

            j += 1

        if company_name:
            entries.append({
                "Company Name": company_name,
                "Address": address,
                "Contact Person": contact_person,
                "Telephone/Mobile": phone,
                "Email": email,
                "Products on Display": products
            })

    i += 1

# Save to Excel
df = pd.DataFrame(entries)
df = df[df["Company Name"].str.strip().astype(bool)]  # remove blanks
df.to_excel(output_excel, index=False)
print(f"✅ Extracted {len(df)} company records. Saved to {output_excel}")
