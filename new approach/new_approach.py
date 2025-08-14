# import re
# import pandas as pd

# # Read the cleaned text file
# with open("aahar2025_cleaned_text.txt", "r", encoding="utf-8") as f:
#     text = f.read()

# # Split each company's details using the separator lines
# companies = re.split(r"-{5,}", text)

# data = []

# for company in companies:
#     company = company.strip()
#     if not company:
#         continue

#     # Extract fields using regex
#     name_match = re.match(r"^(.*?)\n", company)
#     address_match = re.search(r"Address:\s*(.*?)(?=\nContact Person|$)", company, re.S | re.I)
#     contact_person_match = re.search(r"Contact Person\s*:?\s*(.*)", company)
#     email_match = re.search(r"E-?mail\s*:?\s*([^\s,;]+(?:\s*,\s*[^\s,;]+)*)", company, re.I)
#     products_match = re.search(r"Products on Display\s*:?\s*(.*)", company, re.S | re.I)

#     company_name = name_match.group(1).strip() if name_match else ""
#     address = address_match.group(1).replace("\n", " ").strip() if address_match else ""
#     contact_person = contact_person_match.group(1).strip() if contact_person_match else ""
#     email = email_match.group(1).replace("\n", " ").strip() if email_match else ""
#     products = products_match.group(1).replace("\n", " ").strip() if products_match else ""

#     data.append({
#         "Company Name": company_name,
#         "Address": address,
#         "Contact Person": contact_person,
#         "Email": email,
#         "Products on Display": products
#     })

# # Create a DataFrame
# df = pd.DataFrame(data)

# # Save to Excel
# output_file = "aahar2025_extracted.xlsx"
# df.to_excel(output_file, index=False)

# print(f"Data extraction complete. Saved to {output_file}")

import re
import pandas as pd

# Read the cleaned text file
with open("aahar2025_cleaned_text.txt", "r", encoding="utf-8") as f:
    text = f.read()

# Split each company's details using the separator lines
companies = re.split(r"-{5,}", text)

data = []

for company in companies:
    company = company.strip()
    if not company:
        continue

    # Extract fields using regex
    name_match = re.match(r"^(.*?)\n", company)
    address_match = re.search(r"Address:\s*(.*?)(?=\nContact Person|$)", company, re.S | re.I)
    contact_person_match = re.search(r"Contact Person\s*:?\s*(.*)", company)
    email_match = re.search(r"E-?mail\s*:?\s*([^\s,;]+(?:\s*,\s*[^\s,;]+)*)", company, re.I)
    phone_match = re.search(r"Tel\.?/Mobile\s*:?\s*(.*?)(?=\nE-?mail|$)", company, re.S | re.I)
    products_match = re.search(r"Products on Display\s*:?\s*(.*)", company, re.S | re.I)

    # Clean extracted fields
    company_name = name_match.group(1).strip() if name_match else ""
    address = address_match.group(1).replace("\n", " ").strip() if address_match else ""
    contact_person = contact_person_match.group(1).strip() if contact_person_match else ""
    email = email_match.group(1).replace("\n", " ").strip() if email_match else ""
    phone = phone_match.group(1).replace("\n", " ").strip() if phone_match else ""
    products = products_match.group(1).replace("\n", " ").strip() if products_match else ""

    data.append({
        "Company Name": company_name,
        "Address": address,
        "Contact Person": contact_person,
        "Phone": phone,
        "Email": email,
        "Products on Display": products
    })

# Create a DataFrame
df = pd.DataFrame(data)

# Save to Excel
output_file = "aahar2025_extracted2.xlsx"
df.to_excel(output_file, index=False)

print(f"Data extraction complete. Saved to {output_file}")
