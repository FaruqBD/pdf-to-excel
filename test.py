import pandas as pd

# Example data list provided
data_list = [    
    '1 - DRIVING VEHICLE WHILE UNDER THE INFLUENCE OF ALCOHOL', '72F0BC7', 'REED, DYLAN ANTHONY','Citation - Traffic', '07/06/2024', 'Defendant Address:', 'REED, DYLAN ANTHONY', '7634 HANTON AVE', 'SALISBURY, MD 21801', 'Charges:'
]


def extract_case_data(data_list):
    cases = []
    i = 0
    while i < len(data_list):
        if 'Citation - Traffic' in data_list[i]:
            case = {}
            print(i)
            case['Case Number'] = data_list[i - 2] if (i - 2) >= 0 else ''
            case['Style'] = data_list[i - 1] if (i - 1) >= 0 else ''
            case['Defendant Address'] = f"{data_list[i + 3]} {data_list[i + 4]} {data_list[i + 5]}"
            case['Case Type'] = data_list[i]
            case['File Date'] = data_list[i + 1] if (i + 1) < len(data_list) else ''
            cases.append(case)
            i += 7  # Skip the lines related to this case
        else:
            i += 1

    return cases

# Extract case data
case_data = extract_case_data(data_list)

# Create DataFrame
df = pd.DataFrame(case_data)

# Write to Excel
excel_path = 'C:/pyDev/pdf-to-excel/extracted_cases.xlsx'
df.to_excel(excel_path, index=False)

print(f"Data has been written to {excel_path}")
