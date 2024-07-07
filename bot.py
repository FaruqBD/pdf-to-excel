import fitz  # PyMuPDF
import pandas as pd

# Function to extract text from the PDF
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text_data = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        text_data.append(text)
    return text_data

# Function to process the extracted text and write to Excel
def write_selected_data_to_excel(text_data, excel_path):
    selected_data = []

    for text in text_data:
        lines = text.split('\n')
        i = 0
        while i < len(lines):
            if 'Citation - Traffic' in lines[i]:
                case = {}
                if i - 2 >= 0:
                    case['Case Number'] = lines[i - 2].strip()
                else:
                    case['Case Number'] = ''
                
                if i - 1 >= 0:
                    case['Style'] = lines[i - 1].strip()
                else:
                    case['Style'] = ''     
                
                case['Case Type'] = lines[i].strip()
                
                if i + 3 < len(lines) and i + 4 < len(lines) and i + 5 < len(lines):
                    case['Defendant Address'] = f"{lines[i + 3].strip()} {lines[i + 4].strip()} {lines[i + 5].strip()}"
                else:
                    case['Defendant Address'] = ''

                if i + 1 < len(lines):
                    case['File Date'] = lines[i + 1].strip()
                else:
                    case['File Date'] = ''
                
                selected_data.append(case)
                i += 6  # Move past the lines related to this case
            else:
                i += 1

    df = pd.DataFrame(selected_data)
    df.to_excel(excel_path, index=False)

# Paths
pdf_path = 'C:/pyDev/pdf-to-excel/file2024-07-07.pdf'
excel_path = 'C:/pyDev/pdf-to-excel/extracted_cases.xlsx'

# Extract text and write to Excel
text_data = extract_text_from_pdf(pdf_path)
write_selected_data_to_excel(text_data, excel_path)

print(f"Data has been written to {excel_path}")
