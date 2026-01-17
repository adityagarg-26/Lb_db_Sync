import pandas as pd
import re
 
df_master = pd.read_excel(r"C:\Users\Bhuvan\Downloads\lab_tests_export - 2026-01-12T142825.383.xlsx")
df_lab = pd.read_excel(r"C:\Users\Bhuvan\OneDrive - Guraan Tech Pvt. Ltd\Desktop\LOCAL LABS NEW FILES\Balaji Diagno\Verify\Balaji Diagno Pricing Update.xlsx")

# --- Step 1: Direct match on test_name_lab with test ---
merged = pd.merge(
    df_lab,
    df_master[["Test Id","Test Name1","Fasting Required", "Fasting Time", "Report Time", "Detailed Description"]],
    left_on='Master Test Name',
    right_on='Test Name1',
    how='left'
)


merged['Fasting Time(Hours)']=merged['Fasting Time']
merged['Test Name Id']=merged['Test Id']

final_df = merged[[
    "Test Pricing Id",
    "Test Name Id",
    "Test Name",
    "Master Test Name", 
    "Test Type",
    "Status",
    "Fasting Required", 
    "Slot Time (30/60 mins)",
    "Fasting Time(Hours)", 
    "Report Time", 
    "Vendor Price",
    "Vendor Lb Discount",
    "Vendor App Discount",
    "Detailed Description"
]]

final_df = final_df.drop_duplicates(subset=["Master Test Name"], keep="first")

with pd.ExcelWriter(r"C:\Users\Bhuvan\Downloads\Final_war_Lab_Buddy.xlsx", engine="openpyxl") as writer:
    final_df.to_excel(writer, index=False)