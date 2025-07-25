import fitz  # PyMuPDF
import re

# Step 1: Extract text from PDF and save as raw text file
pdf_path = "Aahar 2025 Fair Guide.pdf"
text_output = "aahar2025_raw_text.txt"

with fitz.open(pdf_path) as doc:
    full_text = ""
    for page in doc:
        full_text += page.get_text()

with open(text_output, "w", encoding="utf-8") as f:
    f.write(full_text)

print(f"Raw text extracted to {text_output}")

# Step 2: Clean and structure the raw text into readable blocks
cleaned_output = "aahar2025_cleaned_text.txt"

# Basic cleaning rules:
# - Break after every company record by detecting some pattern like:
#   capitalized company name followed by new lines and contact data
lines = full_text.split('\n')
cleaned_blocks = []
current_block = []

for line in lines:
    line = line.strip()

    # Likely a new company entry
    if re.match(r"^[A-Z0-9][A-Z0-9 .,&'()-]+$", line) and len(line) > 5:
        if current_block:
            cleaned_blocks.append('\n'.join(current_block))
            current_block = []
    current_block.append(line)

# Add the last block
if current_block:
    cleaned_blocks.append('\n'.join(current_block))

# Save cleaned blocks to file
with open(cleaned_output, "w", encoding="utf-8") as f:
    for block in cleaned_blocks:
        f.write(block + "\n" + "-"*80 + "\n")

print(f"Cleaned text saved to {cleaned_output}")
