"""
This script compares contract and invoice data from an Excel file,
calculates fuzzy matching scores between description and part numbers,
and identifies discrepancies. The discrepancies are saved to a new sheet
in the same Excel file.

Author: Carlos Germosen Polanco
Date: 6/19/2024
"""

from difflib import SequenceMatcher
import pandas as pd

# Load the Excel file
FILE_PATH = "Excel Test.xlsx"
contract_df = pd.read_excel(FILE_PATH, sheet_name="Contract", engine="openpyxl")
invoice_df = pd.read_excel(FILE_PATH, sheet_name="Invoice", engine="openpyxl")

# Calculate fuzzy score
def calculate_fuzzy_score(str1: str, str2: str) -> float:
    """
    This function uses the SequenceMatcher from the difflib module to calculate the similarity 
    ratio between two strings. The similarity ratio is a float between 0 and 1, 
    where 1 means the strings are identical.
    
    Args:
        str1 (str): The first string to compare.
        str2 (str): The second string to compare.

    Returns:
        float: The similarity ratio between the two strings.
    """
    return SequenceMatcher(None, str(str1), str(str2)).ratio()

# Initialize the discrepancies DataFrame
discrepancies = pd.DataFrame(
    columns=[
        "Contract Description", 
        "Invoice Description", 
        "Contract PartNumber", 
        "Invoice PartNumber", 
        "Contract Quantity", 
        "Invoice Quantity", 
        "Fuzzy Score"
        ]
    )

# Compare the Contract and Invoice tables
for _, contract_row in contract_df.iterrows():
    for _, invoice_row in invoice_df.iterrows():
        description_score = calculate_fuzzy_score(
            contract_row["Description"],
            invoice_row["Description"]
            )
        partnumber_score = calculate_fuzzy_score(
            str(contract_row["PartNumber"]),
            str(invoice_row["PartNumber"])
            )

        # Compare similar descritions and part numbers
        if (
            description_score > 0.9
            or partnumber_score > 0.9
            ):
            if (
                contract_row["Quantity"] != invoice_row["Quantity"]
                or partnumber_score < 1
                or description_score < 1
            ):

                # Append new row to discrepancies table
                new_row = pd.DataFrame([{
                    "Contract Description":contract_row["Description"],
                    "Invoice Description": invoice_row["Description"],
                    "Contract PartNumber":contract_row["PartNumber"],
                    "Invoice PartNumber": invoice_row["PartNumber"],
                    "Contract Quantity": contract_row["Quantity"],
                    "Invoice Quantity": invoice_row["Quantity"],
                    "Fuzzy Score": (description_score + partnumber_score) / 2
                    }])
                discrepancies = pd.concat([discrepancies, new_row], ignore_index=True)

# Append Totals as a new row
totals_row = pd.DataFrame([{
    "Contract Description": "", 
    "Invoice Description": "", 
    "Contract PartNumber": "", 
    "Invoice PartNumber": "Total: ", 
    "Contract Quantity": discrepancies["Contract Quantity"].sum(), 
    "Invoice Quantity": discrepancies["Invoice Quantity"].sum(), 
    "Fuzzy Score": ""
}])
discrepancies = pd.concat([discrepancies, totals_row], ignore_index = True)

# Write discrepancies table to 'Discrepancies' sheet in excel
with pd.ExcelWriter(FILE_PATH, engine="openpyxl", mode="a", if_sheet_exists='replace') as writer:
    discrepancies.to_excel(writer, sheet_name="Discrepancies", index=False)

print(discrepancies)

print("Discrepancies saved to the 'Discrepancies' sheet.")
