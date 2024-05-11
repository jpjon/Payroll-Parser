import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

def select_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window
    file_path = filedialog.askopenfilename()  # Show the file dialog
    return file_path

def extract_data_from_page(page_text):
    
    # Extract the date
    lines = page_text.split('\n')
    pay_date = None
    for i, line in enumerate(lines):
        if "DATE:" in line and i > 0:
            date_line = lines[i - 1]
            date_match = re.search(r'(\d{2}/\d{2}/\d{4})', date_line)
            if date_match:
                pay_date = date_match.group(1)
                break
            
    #Federal gross pay is incorrect, actual column name is Taxable Comp Before Tax

    # Change it so that if federal gross pay is present, match it
    # If it is not present then use the Current total gross pay amount instead
    def extract_taxable_comp(text):
        # Regex for FEDERAL TAXABLE GROSS PAY
        federal_taxable_regex = r"FEDERAL TAXABLE GROSS PAY (\d+\.\d+)"
        
        # Regex for TOTAL GROSS PAY
        total_gross_pay_regex = r"TOTAL GROSS PAY((?:\s+\d+.\d+)+)"

        # Attempt to find FEDERAL TAXABLE GROSS PAY first
        federal_taxable_match = re.search(federal_taxable_regex, text)
        if federal_taxable_match:
            return federal_taxable_match.group(1)

        # If not found, try to find TOTAL GROSS PAY
        total_gross_pay_match = re.search(total_gross_pay_regex, text)
        if total_gross_pay_match:
            numbers = total_gross_pay_match.group(1).strip().split()
            if len(numbers) >= 2:
                return numbers[-2]

        return None
    
    # Regular expressions for each required field
    name_regex = r"PAY TO THE ([\w\s.]+) AMOUNT:"
    check_no_regex = r"Check Number\s*State:.*?\n(\d+)"
    net_pay_regex = r"NET PAY ([\d,]+\.\d{2})"
    fed_tax_regex = r"Fed Inc Tax (\d+\.\d{2})"
    soc_sec_regex = r"Soc Sec Tax (\d+\.\d{2})"
    # taxable_comp_regex = r"TOTAL GROSS PAY [\d.]+\s([\d,]+\.\d{2})"
    reg_wages_regex = r"Reg wages [\d.]+\s[\d.]+\s([\d,]+\.\d{2})"
    salary_regex = r"Salary\s+(\d+\.\d+)"
    vac_regex = r"Vac HR\s+\d+\.\d+\s+\d+\.\d+\s+(\d+\.\d+)\s+\d+\.\d+"
    ot_regex = r"OT [\d.]+\s[\d.]+\s([\d,]+\.\d{2})"
    holiday_regex = r"HOLIDAY [\d.]+\s[\d.]+\s([\d,]+\.\d{2})"
    sick_regex = r"SICK HOURLY [\d.]+\s[\d.]+\s([\d,]+\.\d{2})"
    bonus_regex = r"Bonus [\d.]+\s[\d.]+\s([\d,]+\.\d{2})"
    scorp_med_regex = r"SCorp Med\s+(\d+\.\d{2})"
    scorp_denta_regex = r"SCorp Denta\s+(\d+\.\d{2})"
    anthem_regex = r"Anthem Med\s+(\d+\.\d+)"
    dental_regex = r"\* DDental\s+(\d+\.\d+)"
    vision_regex = r"Vision \([A-Z]\) (\d+\.\d+)"
    four01k_amount_regex = r"401k\s+(\d+\.00)"
    four01k_percent_regex = r"401k\s+(\d+\.(?!00)\d{2})"
    st_roth_ira_amount_regex = r"St Roth IRA\s+(\d+\.00)"
    st_roth_ira_percent_regex = r"St Roth IRA\s+(\d+\.(?!00)\d{2})"
    reg_hrs_regex = r"Reg wages\s+\d+\.\d+\s+(\d+\.\d+)\s+\d+\.\d+\s+\d+\.\d+(?=\s+[A-Z])"
    ot_hrs = r"OT\s+\d+\.\d+\s+(\d+\.\d+)\s+\d+\.\d+\s+\d+\.\d+"
    holiday_hrs = r"HOLIDAY\s+\d+\.\d+\s+(\d+\.\d+)\s+\d+\.\d+\s+\d+\.\d+"
    bonus_hrs = r"Bonus\s+\d+\.\d+\s+(\d+\.\d+)\s+\d+\.\d+\s+\d+\.\d+"
    sick_hrs = r"SICK HOURLY\s+\d+\.\d+\s+(\d+\.\d+)\s+\d+\.\d+\s+\d+\.\d+"
    vac_hrs = r"Vac HR\s+\d+\.\d+\s+(\d+\.\d+)\s+\d+\.\d+\s+\d+\.\d+"
    reg_rate = r"Reg wages\s+(\d+\.\d+)\s+\d+\.\d+\s+\d+\.\d+\s+\d+\.\d+(?=\s+[A-Z])"
    
    # Extracting each field
    name = re.search(name_regex, page_text)
    check_no = re.search(check_no_regex, page_text)
    net_pay = re.search(net_pay_regex, page_text)
    fed_tax = re.search(fed_tax_regex, page_text)
    soc_sec = re.search(soc_sec_regex, page_text)
    # taxable_comp = re.search(taxable_comp_regex, page_text)
    reg_wages = re.search(reg_wages_regex, page_text)
    salary = re.search(salary_regex, page_text)
    vac = re.search(vac_regex, page_text)
    ot = re.search(ot_regex, page_text)
    holiday = re.search(holiday_regex, page_text)
    sick = re.search(sick_regex, page_text)
    bonus = re.search(bonus_regex, page_text)
    scorp_med = re.search(scorp_med_regex, page_text)
    scrop_denta = re.search(scorp_denta_regex, page_text)
    anthem = re.search(anthem_regex, page_text)
    dental = re.search(dental_regex, page_text)
    vision = re.search(vision_regex, page_text)
    four01k_amount = re.search(four01k_amount_regex, page_text)
    four01k_percent = re.search(four01k_percent_regex, page_text)
    roth_ira_amount = re.search(st_roth_ira_amount_regex, page_text)
    roth_ira_percent = re.search(st_roth_ira_percent_regex, page_text)
    regular_hrs = re.search(reg_hrs_regex, page_text)
    ot_hrs = re.search(ot_hrs, page_text)
    holiday_hrs = re.search(holiday_hrs, page_text)
    bonus_hrs = re.search(bonus_hrs, page_text)
    sick_hrs = re.search(sick_hrs, page_text)
    vac_hrs = re.search(vac_hrs, page_text)
    reg_rate = re.search(reg_rate, page_text)


    # Constructing the data dictionary
    data = {
        'Name': name.group(1) if name else '',
        'Check No': check_no.group(1) if check_no else '',
        'Pay Date': pay_date if pay_date else '',
        'Net': net_pay.group(1) if net_pay else '',
        'Fed Tax W/H': fed_tax.group(1) if fed_tax else '',
        'Soc. Sec. W/H': soc_sec.group(1) if soc_sec else '',
        'Taxable Comp Before Tax': extract_taxable_comp(page_text),
        'Reg wages Before Tax': reg_wages.group(1) if reg_wages else '',
        'Salary Before Tax': salary.group(1) if salary else '',
        'Vac HR Before Tax': vac.group(1) if vac else '',
        'OT Before Tax': ot.group(1) if ot else '',
        'HOLIDAY HR Before Tax': holiday.group(1) if holiday else '',
        'SICK HR Before Tax': sick.group(1) if sick else '',
        'Bonus Before Tax': bonus.group(1) if bonus else '',
        'SCorp Med  W/H' : scorp_med.group(1) if scorp_med else '',
        'SCorp DI' : scrop_denta.group(1) if scrop_denta else '',
        'UHC / Anthem Med W/H' : anthem.group(1) if anthem else '',
        "Ddental W/H" : dental.group(1) if dental else '',
        "Vision  W/H + O" : vision.group(1) if vision else '',
        "401K by Amt" : four01k_amount.group(1) if four01k_amount else '',
        "401k by %" : four01k_percent.group(1) if four01k_percent else '',
        "ROTH IRA by Amt" : roth_ira_amount.group(1) if roth_ira_amount else '',
        "ROTH IRA by %" :  roth_ira_percent.group(1) if roth_ira_percent else '',
        "Rate" : reg_rate.group(1) if reg_rate else '',
        "Regular Hr" :  regular_hrs.group(1) if regular_hrs else '',
        "Ot Hr" :  ot_hrs.group(1) if ot_hrs else '',
        "Holiday Hr" :  holiday_hrs.group(1) if holiday_hrs else '',
        "Bonus Hr" :  bonus_hrs.group(1) if bonus_hrs else '',
        "Sick Hr" :  sick_hrs.group(1) if sick_hrs else '',
        "Vacation Hr" :  vac_hrs.group(1) if vac_hrs else ''
    }

    return data

def extract_payroll_data(pdf_text):
    # Split the text by 'NEXT PAGE' and process each page
    pages = pdf_text.split('NEXT PAGE')
    extracted_data = []
    for page in pages:
        data = extract_data_from_page(page)
        extracted_data.append(data)
    return extracted_data

def main():
    file_path = select_file()  # Get file path using file dialog
    if not file_path:
        print("No file selected. Exiting...")
        return
    
    with pdfplumber.open(file_path) as pdf:
        all_text = ''
        for page in pdf.pages:
            text = page.extract_text()
            all_text += text + '\nNEXT PAGE\n'

    payroll_data = extract_payroll_data(all_text)

    # Create DataFrame and save to Excel
    columns = ['Name', 'Check No', 'Pay Date', 'Net', 'Fed Tax W/H', 'Soc. Sec. W/H',
               'Taxable Comp Before Tax', 'Reg wages Before Tax', 'Salary Before Tax',
               'Vac HR Before Tax', 'OT Before Tax', 'HOLIDAY HR Before Tax',
               'SICK HR Before Tax', 'Bonus Before Tax', 'SCorp Med  W/H', 'SCorp DI', 'UHC / Anthem Med W/H', "Ddental W/H", 
               "Vision  W/H + O", "401K by Amt", "401k by %", "ROTH IRA by Amt", "ROTH IRA by %", "Rate", "Regular Hr", "Ot Hr", "Holiday Hr", "Bonus Hr",
                "Sick Hr", "Vacation Hr"]
    df = pd.DataFrame(payroll_data, columns=columns)

     # Converting numeric columns to floats
    numeric_cols = ['Net', 'Fed Tax W/H', 'Soc. Sec. W/H', 'Taxable Comp Before Tax', 
                    'Reg wages Before Tax', 'Salary Before Tax', 
                    'Vac HR Before Tax', 'OT Before Tax', 'HOLIDAY HR Before Tax', 
                    'SICK HR Before Tax', 'Bonus Before Tax','SCorp Med  W/H', 'SCorp DI', 'UHC / Anthem Med W/H', "Ddental W/H", 
                    "Vision  W/H + O", "401K by Amt", "401k by %", "ROTH IRA by Amt", "ROTH IRA by %", "Rate", "Regular Hr", "Ot Hr", "Holiday Hr", "Bonus Hr",
                    "Sick Hr", "Vacation Hr"]
    for col in numeric_cols:
        if df[col].notnull().any():  # Check if the column is not completely empty
            df[col] = df[col].replace({',': ''}, regex=True).apply(pd.to_numeric, errors='coerce')

    # Get the current date
    current_date = datetime.now()

    # Format the date as MM_DD_YYYY
    formatted_date = current_date.strftime("%m_%d_%Y")

    # Construct the filename
    filename = f"Payroll_{formatted_date}.xlsx"

    # Save the DataFrame to the Excel file
    df.to_excel(filename, index=False)


if __name__ == "__main__":
    main()
