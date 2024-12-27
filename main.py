import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
import us
import os
import config  

input_file = config.input_file
output_file = config.output_file
target_date = config.target_date


df = pd.read_excel(input_file, sheet_name="Sheet1")


df.columns = df.columns.str.strip()

# Filter rows based on the target start date
df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")  
df_selected = df[df["Start Date"] == pd.to_datetime(target_date)]  


def get_state_abbreviation(state_name):
    state_abbr = us.states.lookup(state_name)
    return state_abbr.abbr if state_abbr else state_name.upper()[:2]  


package_descriptions = ["laptop", "monitor", "WFH Bundle"]


output_data = pd.DataFrame()

# Repeat each row three times with different item descriptions
for _, row in df_selected.iterrows():
    for item_description in package_descriptions:
        package_type = "FEDEX_EXTRA_LARGE_BOX" if item_description == "monitor" else "FEDEX_LARGE_BOX"
        ship_date = (row["Start Date"] - pd.Timedelta(days=6)).strftime("%Y%m%d")  

        if item_description == "monitor":
            package_weight = 8
        elif item_description == "WFH Bundle":
            package_weight = 4
        elif item_description == "laptop":
            package_weight = 5
        else:
            package_weight = ""

        output_data = pd.concat(
            [
                output_data,
                pd.DataFrame({
                    "senderContactName": ["Better - IT Department"],
                    "senderCompany": ["Better"],
                    "senderContactNumber": ["9294273517"],
                    "senderEmail": ["clei@better.com"],
                    "senderLine1": ["59 beach street FL 3"],
                    "senderPostcode": ["10013"],
                    "senderState": ["NY"],
                    "senderCity": ["NYC"],
                    "senderCountry": ["US"],
                    "recipientContactName": [row["Full Name (First Name, Last Name)"]],
                    "recipientLine1": [row["Street Address (including apartment/unit number, if applicable)"]],
                    "recipientCity": [row["City"]],
                    "recipientState": [get_state_abbreviation(row["State"])],
                    "recipientPostcode": [row["ZIP Code"]],
                    "recipientEmail": [row["Email Address"]],
                    "recipientContactNumber": [row["Phone Number"]],
                    "recipientCountry": ["US"],
                    "packageType": [package_type],
                    "numberOfPackages": [1],
                    "packageWeight": [package_weight],  
                    "weightUnits": ["LBS"],
                    "itemDescription": [item_description],
                    "currencyType": ["USD"],
                    "serviceType": ["FEDEX_2_DAY"],
                    "signatureType": ["NO_SIGNATURE_REQUIRED"],
                    "recipientDeliveryNotification": ["N"],
                    "recipientShipAlertNotification": ["N"],
                    "recipientExceptionNotification": ["N"],
                    "shipDate": [ship_date],  
                    "oneRatePricing": ["Y"],
                })
            ],
            ignore_index=True,
        )


sheet_name = pd.to_datetime(target_date).strftime("%Y-%m-%d")  

# Write the new DataFrame to an Excel file
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    output_data.to_excel(writer, sheet_name=sheet_name, index=False)

  
    workbook = writer.book
    sheet = writer.sheets[sheet_name]

    
    border_style = Border(
        left=Side(border_style="thin"),
        right=Side(border_style="thin"),
        top=Side(border_style="thin"),
        bottom=Side(border_style="thin"),
    )
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    font = Font(size=12)

    for row in sheet.iter_rows():
        for cell in row:
            cell.border = border_style
            cell.alignment = alignment
            cell.font = font

    
    for row in sheet.iter_rows():
        sheet.row_dimensions[row[0].row].height = 40

    for col in sheet.columns:
        max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        sheet.column_dimensions[col[0].column_letter].width = max_length + 2


if os.name == 'nt':  
    os.startfile(output_file)
elif os.name == 'posix':  
    os.system(f"open {output_file}")
else:
    print(f"File created: {output_file}. Please open it manually.")

print(f"New Excel file '{output_file}' created with the sheet name '{sheet_name}' and opened successfully!")
