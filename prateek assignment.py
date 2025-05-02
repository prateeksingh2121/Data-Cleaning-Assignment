#!/usr/bin/env python
# coding: utf-8

# In[4]:


pip install pdfplumber pandas openpyxl


# In[12]:


import zipfile

import os

zip_path= "D:\Data analyst\Mann Hospitality LLP\Cleaning of Data & Merging into single excel.zip"
extract_to= "D:\Data analyst\Mann Hospitality LLP\Assignment"

with zipfile.ZipFile(zip_path,'r') as zip_ref:
    zip_ref.extractall(extract_to)
    
print("Zip Extracted successfully.")


# In[1]:


import pandas as pd
import os
from openpyxl import load_workbook

input_folder=r"D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Payout Summary & Order Level Sales"
output_file="D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Data Analyst Quick Assignment.xlsx"
target_sheet="Summary"

summary_records=[]

for file_name in os.listdir(input_folder):
    if file_name.endswith(".xlsx"):
        file_path= os.path.join(input_folder,file_name)
        try:
            df=pd.read_excel(file_path, sheet_name="Summary", header=None)
            summary_data = {
                "Brand": df.iloc[4, 1],
                "Location": df.iloc[5, 1],
                "City": df.iloc[6, 1],
                "Res-Id": df.iloc[7, 1],
                "Payout Period": df.iloc[11, 2],
                "Payout Settlement Date": df.iloc[12, 2],
                "Total Payout": df.iloc[13, 2],
                "Total Orders (Delivered + Cancelled)": df.iloc[14, 2],
                "Bank UTR": df.iloc[15, 2],
                "File Name": file_name
            }
            summary_records.append(summary_data)
        except Exception as e:
            print(f"Error in {file_name}:{e}")

summary_df= pd.DataFrame(summary_records)

with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:

 summary_df.to_excel(writer, sheet_name=target_sheet, startrow=1, index=False, header=False)

print("Data written to target sheet starting below header.")


# In[29]:


import pandas as pd
import os

input_folder = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Payout Summary & Order Level Sales"
output_file = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Data Analyst Quick Assignment.xlsx"
target_sheet = "Payout Breakup Tab"

payout_breakup_records = []

for file_name in os.listdir(input_folder):
    if not file_name.endswith(".xlsx") or file_name.startswith("~$"):
        continue

    file_path = os.path.join(input_folder, file_name)
    try:
     
        df_payout = pd.read_excel(file_path, sheet_name="Payout Breakup", header=None)
        
       
        df_summary = pd.read_excel(file_path, sheet_name="Summary", header=None)
        brand = df_summary.iloc[4, 1] if not pd.isna(df_summary.iloc[4, 1]) else ''
        res_id = df_summary.iloc[7, 1] if not pd.isna(df_summary.iloc[7, 1]) else ''
        payout_period = df_summary.iloc[11, 2] if not pd.isna(df_summary.iloc[11, 2]) else ''

        sr_no = 1
        for row_idx in range(2, len(df_payout)): 
            try:
                particular = df_payout.iloc[row_idx, 2] 
                delivered = df_payout.iloc[row_idx, 3]   
                cancelled = df_payout.iloc[row_idx, 4]   
                total = df_payout.iloc[row_idx, 5]       

                if pd.isna(particular):
                    break  

                payout_breakup_records.append({
                    "SR.No.": sr_no,
                    "Particulars": particular,
                    "Delivered Orders": delivered,
                    "Cancelled Orders": cancelled,
                    "Total": total,
                    "Brand": brand,
                    "Res-Id": res_id,
                    "Payout Period": payout_period,
                    "File Name": file_name
                })
                sr_no += 1
            except Exception as e:
                print(f"Error parsing row {row_idx + 1} in {file_name}: {e}")

    except Exception as e:
        print(f"Error processing {file_name}: {e}")

if payout_breakup_records:
    payout_df = pd.DataFrame(payout_breakup_records)
    payout_df["Cancelled Orders"] = pd.to_numeric(payout_df["Cancelled Orders"], errors="coerce").fillna(0).astype(int)
    payout_df["Total"] = pd.to_numeric(payout_df["Total"], errors="coerce").fillna(0).astype(int)

    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        payout_df.to_excel(writer, sheet_name=target_sheet, startrow=1, index=False, header=False)

    print("Payout Breakup data written to Excel.")
else:
    print("No records found to write.")


# In[17]:


import pandas as pd
import os

input_folder = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Payout Summary & Order Level Sales"
output_file = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Data Analyst Quick Assignment.xlsx"
target_sheet = "Order Level"
summary_sheet = "Summary"

summary_df = pd.read_excel(output_file, sheet_name=summary_sheet)
summary_lookup = summary_df.set_index("File Name")[["Brand", "Res-Id", "Payout Period"]].to_dict("index")

order_level_data = []

for file_name in os.listdir(input_folder):
    if file_name.endswith(".xlsx") and not file_name.startswith("~$"):
        file_path = os.path.join(input_folder, file_name)
        try:
            df = pd.read_excel(file_path, sheet_name="Order Level", header=None)

            header_row_idx = None
            for i in range(min(15, len(df))):
                row = df.iloc[i, :5].astype(str).str.lower()
                if any(cell.startswith("sr") or "order id" in cell for cell in row):
                    header_row_idx = i
                    break

            if header_row_idx is None:
                print(f"Skipping {file_name} — header row not found.")
                continue

            data_df = df.iloc[header_row_idx + 1:].copy()
            data_df.columns = df.iloc[header_row_idx]
            data_df = data_df.reset_index(drop=True)
            data_df = data_df[data_df.iloc[:, 0].notna()] 

            meta = summary_lookup.get(file_name, {})
            brand = meta.get("Brand", "NA")
            res_id = meta.get("Res-Id", "NA")
            payout_period = meta.get("Payout Period", "NA")

            data_df["Brand"] = brand
            data_df["Res-Id"] = res_id
            data_df["Payout Period"] = payout_period
            data_df["File Name"] = file_name

            order_level_data.append(data_df)

        except Exception as e:
            print(f"Error processing {file_name}: {e}")

if order_level_data:
    final_df = pd.concat(order_level_data, ignore_index=True)

    template = pd.read_excel(output_file, sheet_name=target_sheet, nrows=0)
    final_df = final_df.reindex(columns=template.columns)

    final_df.fillna("NA", inplace=True)

    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name=target_sheet, index=False)

    print("Order Level data written successfully.")
else:
    print("No Order Level data found.")


# In[15]:


import os
import pdfplumber
import pandas as pd
import re

# === 1. File paths ===
input_folder = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Commission Invoices"  # Folder with all PDF files
output_file = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Combined Commission Invoice\Commission Invoice.xlsx"  # Output Excel file

records = []

def extract_field(text, label, fallback="NA"):
    match = re.search(rf"{label}\s*[:\-]?\s*([^\n]+)", text, re.IGNORECASE)
    return match.group(1).strip() if match else fallback

for filename in os.listdir(input_folder):
    if filename.endswith(".pdf"):
        filepath = os.path.join(input_folder, filename)

        try:
            with pdfplumber.open(filepath) as pdf:
                full_text = "\n".join(p.extract_text() for p in pdf.pages if p.extract_text())
            lines = full_text.splitlines()

            invoice_number = extract_field(full_text, r"Invoice Number")
            invoice_date = extract_field(full_text, r"Invoice Date")
            pan = extract_field(full_text, r"PAN")
            swiggy_gstin = extract_field(full_text, r"GSTIN")
            brand_id = extract_field(full_text, r"Restaurant / Store ID")
            brand_name = extract_field(full_text, r"Restaurant / Store Name")
            restaurant_gstin = extract_field(full_text, r"Restaurant GSTIN")
            payout_period = extract_field(full_text, r"Service Period")
            irn = extract_field(full_text, r"IRN")
            fy_year = "2024-25"
            year = "2025"
            month = "April"

            found = False
            for line in lines:
                if "Service Fee" in line and "996211" in line:
                    parts = line.strip().split()
                    try:
                        description = "Service Fee"
                        hsn = "996211"
                        unit = "OTH"
                        quantity = 1
                        unit_price = parts[-6].replace(",", "")
                        base_amount = parts[-5].replace(",", "")
                        cgst_amount = parts[-4].replace(",", "")
                        sgst_amount = parts[-3].replace(",", "")
                        total_amount = parts[-1].replace(",", "")
                        found = True
                    except:
                        unit_price = base_amount = cgst_amount = sgst_amount = total_amount = "0"

                    records.append({
                        "payout_period": payout_period,
                        "file_name": filename,
                        "brand_id": brand_name,
                        "res_id": brand_id,
                        "pan": pan,
                        "gstin": restaurant_gstin,
                        "mann_gstin": "NA",
                        "swiggy_gstin": swiggy_gstin,
                        "fy_year": fy_year,
                        "year": year,
                        "month": month,
                        "invoice_date": invoice_date,
                        "invoice_number": invoice_number,
                        "original_invoice_number": "NA",
                        "invoice_type": "Regular",
                        "sr_no": 1,
                        "description": description,
                        "hsn": hsn,
                        "unit_of_measure": unit,
                        "quantity": quantity,
                        "unit_price": unit_price,
                        "base_amount": base_amount,
                        "discount": 0,
                        "assessable_value": base_amount,
                        "cgst_rate": 9,
                        "cgst_amount": cgst_amount,
                        "sgst_rate": 9,
                        "sgst_amount": sgst_amount,
                        "igst_rate": 0,
                        "igst_amount": 0,
                        "total_amount": total_amount,
                        "grand_total": total_amount,
                        "other_charges_reimbursement_of_discount": "0",
                        "irn": irn
                    })
                    break

            if not found:
                print(f"No service line found in: {filename}")

        except Exception as e:
            print(f"Error processing {filename}: {e}")

if records:
    df = pd.DataFrame(records)
    df.to_excel(output_file, index=False)
    print(f"\Extracted to:{output_file}")
else:
    print("No data extracted from any PDF.")


# In[25]:


import pandas as pd

commission_data_path = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Combined Commission Invoice\Commission Invoice.xlsx"
assignment_path = r"D:\Data analyst\Mann Hospitality LLP\Assignment\Cleaning of Data & Merging into single excel\Data Analyst Quick Assignment.xlsx"
target_sheet = "Commission Invoice"

commission_df = pd.read_excel(commission_data_path)

assignment_columns = pd.read_excel(assignment_path, sheet_name=target_sheet, nrows=0).columns.tolist()

for col in assignment_columns:
    if col not in commission_df.columns:
        if col == "mann_gstin":
            commission_df[col] = "NA"
        else:
            commission_df[col] = 0

if "mann_gstin" in commission_df.columns:
    commission_df["mann_gstin"] = commission_df["mann_gstin"].fillna("NA")

commission_df = commission_df[assignment_columns]

with pd.ExcelWriter(assignment_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    commission_df.to_excel(writer, sheet_name=target_sheet, index=False)

print("Commission Invoice sheet updated with all required data (mann_gstin filled, missing values = 0).")


# In[ ]:





# In[ ]:





# In[ ]:




