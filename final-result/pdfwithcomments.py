# Required Libraries
import pandas as pd                      # For data manipulation and saving to Excel
import PyPDF2                            # (Not used in current code – can be removed)
import pdfplumber                        # To extract text from PDF pages accurately
import re                                # For regular expressions to match structured patterns
import openpyxl                          # Used by pandas to write Excel files
from pathlib import Path                 # For file path and existence checking
import logging                           # For logging progress and errors

# Setup basic logging configuration
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Main class that handles data extraction from the PDF
class CompanyDataExtractor:
    def __init__(self):
        self.companies = []  # This will hold all extracted company information

    # Function to determine if a page is irrelevant (e.g., advertisements, indexes)
    def is_unwanted_page(self, text):
        # List of patterns that identify unwanted pages
        unwanted_patterns = [
            r'AAHAR.*THE INTERNATIONAL FOOD & HOSPITALITY FAIR',
            r'GOVT\.\s*PARTICIPANTS\s*$',
            r'www\.indiatradefair\.com\s*$',
            r'ORGANISER.*ITPO',
            r'MARCH\s*\|\s*4-8\s*\|\s*2025.*Bharat Mandapam',
            r'^\s*S\.NO\.\s*COMPANY NAME\s*STALL NO\.\s*HALL NO\.\s*$',
            r'^\s*\d+\s+[A-Z\s]+\s+[A-Z0-9\-]+\s+\d+\s*(FF|GF)\s*\n\s*\d+\s+[A-Z\s]+\s+[A-Z0-9\-]+\s+\d+\s*(FF|GF)\s*'
        ]
        
        matches = 0
        for pattern in unwanted_patterns:
            if re.search(pattern, text, re.IGNORECASE | re.MULTILINE):
                matches += 1
        return matches >= 2  # Skip the page only if multiple patterns match (to avoid false positives)

    # Extract all relevant company blocks from full text
    def extract_companies_from_text(self, text):
        companies = []

        # Remove page markers from text
        text = re.sub(r'--- PAGE \d+ ---', '', text)

        # Try multiple ways to split text into company blocks
        company_blocks = []

        # Match patterns that look like company names and grab associated block
        company_pattern = r'([A-Z][A-Z\s&\(\)\.,-]{10,}(?:LIMITED|LTD|PRIVATE|PVT|CORPORATION|ENTERPRISES|INDUSTRIES|AUTHORITY|BUREAU|COMPANY|CO\.|INC|GROUP|ASSOCIATION).*?)(?=\n[A-Z][A-Z\s&\(\)\.,-]{10,}(?:LIMITED|LTD|PRIVATE|PVT|CORPORATION|ENTERPRISES|INDUSTRIES|AUTHORITY|BUREAU|COMPANY|CO\.|INC|GROUP|ASSOCIATION)|$)'
        potential_blocks = re.findall(company_pattern, text, re.DOTALL | re.IGNORECASE)

        # Also try splitting by Hall/STALL format
        hall_stall_splits = re.split(r'HALL\s*:\s*\d+\s*[A-Z]*\s*STALL\s*:\s*[A-Z0-9\-]+', text, flags=re.IGNORECASE)

        # Try splitting by address field
        address_splits = re.split(r'\n(?=Address\s*:)', text, flags=re.IGNORECASE)

        # Combine all extracted blocks
        all_blocks = potential_blocks + hall_stall_splits + address_splits

        for block in all_blocks:
            if not block or len(block.strip()) < 50:  # Skip too-short garbage blocks
                continue

            # Check if the block has required structure
            if self.has_company_structure(block):
                company_data = self.parse_company_block(block)

                # Only keep entries with valid company name
                if company_data and company_data.get('company_name') and len(company_data['company_name']) > 3:
                    companies.append(company_data)

        return companies

    # Check if the text block looks like a valid company entry
    def has_company_structure(self, block):
        required_patterns = [r'Address\s*:', r'Contact Person\s*:']
        optional_patterns = [r'Tel\.?/Mobile\s*:', r'E-mail\s*:', r'Products on Display\s*:', r'HALL\s*:\s*\d+', r'STALL\s*:\s*[A-Z0-9\-]+']

        # Check for required field matches
        required_count = sum(1 for pattern in required_patterns if re.search(pattern, block, re.IGNORECASE))

        # Check for optional field matches
        optional_count = sum(1 for pattern in optional_patterns if re.search(pattern, block, re.IGNORECASE))

        # Valid if 2 required fields, or 1 required + 2 optional
        return required_count >= 2 or (required_count >= 1 and optional_count >= 2)

    # Parse a block of text and extract structured fields into a dictionary
    def parse_company_block(self, block):
        company_data = {
            'company_name': '',
            'address': '',
            'contact_person': '',
            'phone': '',
            'email': '',
            'products': ''
        }

        # Clean and split block into individual lines
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        i = 0

        # Walk through lines one by one
        while i < len(lines):
            line = lines[i]

            # First, try to detect and store company name
            if not company_data['company_name'] and self.looks_like_company_name(line):
                clean_name = re.sub(r'\s*HALL\s*:\s*.*$', '', line, flags=re.IGNORECASE)
                clean_name = re.sub(r'\s*STALL\s*:\s*.*$', '', clean_name, flags=re.IGNORECASE)
                company_data['company_name'] = clean_name.strip()
                i += 1
                continue

            # Extract and join address lines
            address_match = re.search(r'Address\s*:\s*(.+)', line, re.IGNORECASE)
            if address_match:
                address = address_match.group(1).strip()
                i += 1
                while i < len(lines) and not re.search(r'Contact Person|Tel|Mobile|E-mail|Products|HALL|STALL', lines[i], re.IGNORECASE):
                    address += ' ' + lines[i].strip()
                    i += 1
                company_data['address'] = address
                continue

            # Extract Contact Person
            contact_match = re.search(r'Contact Person\s*:\s*(.+)', line, re.IGNORECASE)
            if contact_match:
                company_data['contact_person'] = contact_match.group(1).strip()
                i += 1
                continue

            # Extract phone number
            phone_match = re.search(r'Tel\.?/?Mobile\s*:\s*(.+)', line, re.IGNORECASE)
            if phone_match:
                phone_text = phone_match.group(1).strip()
                phone_clean = re.sub(r'[^\d,\+\-\s\(\)]', ' ', phone_text)
                phone_clean = re.sub(r'\s+', ' ', phone_clean).strip()
                company_data['phone'] = phone_clean
                i += 1
                continue

            # Extract email
            email_match = re.search(r'E-mail\s*:\s*(.+)', line, re.IGNORECASE)
            if email_match:
                email = email_match.group(1).strip()
                email_clean = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', email)
                company_data['email'] = email_clean.group(0) if email_clean else email
                i += 1
                continue

            # Extract products on display
            products_match = re.search(r'Products on Display\s*:\s*(.+)', line, re.IGNORECASE)
            if products_match:
                products = products_match.group(1).strip()
                i += 1
                while i < len(lines) and not re.search(r'Contact Person|Tel|Mobile|E-mail|Address|HALL|STALL|^[A-Z][A-Z\s]{20,}', lines[i], re.IGNORECASE):
                    products += ' ' + lines[i].strip()
                    i += 1
                company_data['products'] = products
                continue

            i += 1  # move to next line

        # Final cleanup: remove extra whitespace
        for key in company_data:
            if company_data[key]:
                company_data[key] = re.sub(r'\s+', ' ', str(company_data[key])).strip()

        return company_data

    # Determines if a line looks like a company name
    def looks_like_company_name(self, line):
        # Skip lines that match known labels like "Address:", "E-mail:", etc.
        skip_patterns = [
            r'^Address\s*:', r'^Contact Person\s*:', r'^Tel\.?/?Mobile\s*:', r'^E-mail\s*:', r'^Products on Display\s*:',
            r'^\d+$', r'^HALL\s*:', r'^STALL\s*:', r'^\d+\s+[A-Z\s]+\s+[A-Z0-9\-]+\s+\d+\s*(FF|GF)\s*$'
        ]
        for pattern in skip_patterns:
            if re.search(pattern, line, re.IGNORECASE):
                return False

        # Company-related keywords
        company_indicators = [
            r'LIMITED|LTD|PRIVATE|PVT|CORPORATION|ENTERPRISES|INDUSTRIES',
            r'AUTHORITY|BUREAU|COMPANY|CO\.|INC|GROUP|ASSOCIATION',
            r'FOODS?|SPICES?|BEVERAGES?|TRADING|EXPORTS?|IMPORTS?'
        ]
        for pattern in company_indicators:
            if re.search(pattern, line, re.IGNORECASE):
                return True

        # If line has 2+ capitalized words and is long enough, consider it a name
        if len(line) > 10:
            words = line.split()
            capital_words = sum(1 for word in words if word and word[0].isupper())
            if capital_words >= 2 and len(words) >= 2:
                return True

        return False

    # Extracts text from a PDF using pdfplumber
    def extract_text_pdfplumber(self, pdf_path):
        try:
            full_text = ""
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text()
                    if page_text:
                        logger.info(f"Processing page {page_num}")
                        if not self.is_unwanted_page(page_text):
                            full_text += f"\n--- PAGE {page_num} ---\n" + page_text
                        else:
                            logger.info(f"Skipping page {page_num}")
            return full_text
        except Exception as e:
            logger.error(f"Error with pdfplumber: {e}")
            return ""

    # Extract company info from PDF file path
    def extract_companies_from_pdf(self, pdf_path):
        logger.info(f"Processing PDF: {pdf_path}")
        text = self.extract_text_pdfplumber(pdf_path)
        if not text:
            logger.error("No text extracted from PDF")
            return []

        companies = self.extract_companies_from_text(text)

        # Remove duplicates by company name
        unique_companies = []
        seen_names = set()

        for company in companies:
            name = company.get('company_name', '').strip().upper()
            if name and name not in seen_names:
                seen_names.add(name)
                unique_companies.append(company)

        logger.info(f"Extracted {len(unique_companies)} unique companies")
        return unique_companies

    # Save extracted companies to Excel
    def save_to_excel(self, companies, output_path):
        if not companies:
            logger.warning("No companies to save")
            return

        df = pd.DataFrame(companies)

        # Set desired column order and rename headers for readability
        column_order = ['company_name', 'address', 'contact_person', 'phone', 'email', 'products']
        df = df.reindex(columns=column_order)
        df.columns = ['Company Name', 'Address', 'Contact Person', 'Phone Number', 'Email', 'Products on Display']

        # Replace empty or NaN with "N/A"
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace('nan', '')
            df[col] = df[col].replace('', 'N/A')

        # Write to Excel with adjusted column widths
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Company_Data', index=False)
            worksheet = writer.sheets['Company_Data']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        logger.info(f"Company data saved to {output_path}")
        return df

# Main script that runs the extraction process
def main():
    PDF_PATH = "D:/10xtech/projects/pdf extraction/Aahar 2025 Fair Guide.pdf"
    OUTPUT_PATH = "company_data.xlsx"

    # Ensure the PDF exists
    if not Path(PDF_PATH).exists():
        logger.error(f"PDF file not found: {PDF_PATH}")
        return

    extractor = CompanyDataExtractor()

    try:
        # Extract company data
        companies = extractor.extract_companies_from_pdf(PDF_PATH)

        if companies:
            # Save to Excel
            df = extractor.save_to_excel(companies, OUTPUT_PATH)
            print(f"Successfully extracted {len(companies)} companies from {PDF_PATH}")
            print(f"Company data saved to {OUTPUT_PATH}")

            # Show sample
            print("\nSample of Extracted Data:")
            sample = companies[0]
            for key, value in sample.items():
                if value and value != 'N/A':
                    print(f"{key.replace('_', ' ').title()}: {value[:100]}{'...' if len(str(value)) > 100 else ''}")

            # Summary Stats
            print("\nExtraction Summary:")
            print(f"Total companies extracted: {len(companies)}")
            print(f"Companies with addresses: {sum(1 for c in companies if c.get('address') and c.get('address') != 'N/A')}")
            print(f"Companies with contact persons: {sum(1 for c in companies if c.get('contact_person') and c.get('contact_person') != 'N/A')}")
            print(f"Companies with phone numbers: {sum(1 for c in companies if c.get('phone') and c.get('phone') != 'N/A')}")
            print(f"Companies with emails: {sum(1 for c in companies if c.get('email') and c.get('email') != 'N/A')}")
        else:
            print("No company data found in the PDF")

    except Exception as e:
        logger.error(f"Error processing PDF: {e}")

# Run script
if __name__ == "__main__":
    main()
