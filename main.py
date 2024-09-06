import pandas as pd
import numpy as np

# Load your Excel file
file_path = 'KidStationsDocs.xlsx'  # replace with your file path
data = pd.read_excel(file_path, header=1)

# Identify the relevant columns
doc_a_column = 'A Statement'  # Replace with the actual column name if different
doc_b_column = 'B Statement'  # Replace with the actual column name for Doc B
doc_c_column = 'C Statement'  # Replace with the actual column name for Doc C

district_column = 'DistrictName'
progress_code_column = 'Main_Progress_Code'  # Replace with the actual column name if different
landowner_column = 'Landowner'  # Replace with the actual column name for Landowner
center_name_column = 'StationName'  # Replace with the actual column name for the center name

# Replace missing values in the 'A Statement', 'B Statement', 'C Statement', and 'Main_Progress_Code' columns with 0
data[doc_a_column] = data[doc_a_column].fillna(0).astype(int)
data[doc_b_column] = data[doc_b_column].fillna(0).astype(int)
data[doc_c_column] = data[doc_c_column].fillna(0).astype(int)
data[progress_code_column] = data[progress_code_column].fillna(0).astype(int)

# Translate progress codes
progress_code_mapping = {
    0: 'ไม่ยอม',
    1: 'ไม่มีอะไรเลย / ไม่รู้',
    2: 'ยอม แต่ไม่มีลายลักอักษร',
    3: 'ผ่าน',
    4: 'ยอมแต่เอกสารไม่ครบ',
    5: 'ระหว่างดำเนินการขอเอกสาร',
    6: 'รอทางเขตติดต่อกลับ'
}

# Apply the translation to the progress codes
data['Translated_Progress_Code'] = data[progress_code_column].map(progress_code_mapping)

# Create columns indicating whether Doc A, B, and C are present
data['Has_Doc_A'] = data[doc_a_column].apply(lambda x: 'Yes' if x == 1 else 'No')
data['Has_Doc_B'] = data[doc_b_column].apply(lambda x: 'Yes' if x == 1 else 'No')
data['Has_Doc_C'] = data[doc_c_column].apply(lambda x: 'Yes' if x == 1 else 'No')

# Ensure that the grouping keeps all districts for Doc A, B, and C
doc_summary = data.groupby(district_column, as_index=False).apply(
    lambda df: pd.Series({
        'Without_Doc_A': (df[doc_a_column] == 0).sum(),
        'With_Doc_A': (df[doc_a_column] == 1).sum(),
        'Without_Doc_B': (df[doc_b_column] == 0).sum(),
        'With_Doc_B': (df[doc_b_column] == 1).sum(),
        'Without_Doc_C': (df[doc_c_column] == 0).sum(),
        'With_Doc_C': (df[doc_c_column] == 1).sum(),
        'Total_Centers': len(df)
    })
).reset_index(drop=True)

# Ensure the 'Landowner_Type' column is correctly created
data['Landowner_Type'] = data[landowner_column]  # Simply use the existing Landowner column

# Reuse the calculations for analysis by Landowner_Type
landowner_doc_summary = data.groupby('Landowner_Type', as_index=False).apply(
    lambda df: pd.Series({
        'Without_Doc_A': (df[doc_a_column] == 0).sum(),
        'With_Doc_A': (df[doc_a_column] == 1).sum(),
        'Without_Doc_B': (df[doc_b_column] == 0).sum(),
        'With_Doc_B': (df[doc_b_column] == 1).sum(),
        'Without_Doc_C': (df[doc_c_column] == 0).sum(),
        'With_Doc_C': (df[doc_c_column] == 1).sum(),
        'Total_Centers': len(df)
    })
).reset_index(drop=True)

# Calculate the percentage of centers with each document in each district
doc_summary['Percentage_With_Doc_A'] = (doc_summary['With_Doc_A'] / doc_summary['Total_Centers']) * 100
doc_summary['Percentage_With_Doc_B'] = (doc_summary['With_Doc_B'] / doc_summary['Total_Centers']) * 100
doc_summary['Percentage_With_Doc_C'] = (doc_summary['With_Doc_C'] / doc_summary['Total_Centers']) * 100

# Calculate the percentage of centers with each document for each Landowner type
landowner_doc_summary['Percentage_With_Doc_A'] = (landowner_doc_summary['With_Doc_A'] / landowner_doc_summary['Total_Centers']) * 100
landowner_doc_summary['Percentage_With_Doc_B'] = (landowner_doc_summary['With_Doc_B'] / landowner_doc_summary['Total_Centers']) * 100
landowner_doc_summary['Percentage_With_Doc_C'] = (landowner_doc_summary['With_Doc_C'] / landowner_doc_summary['Total_Centers']) * 100

# Analysis of Main Progress Code with swapped axis
progress_summary = data.groupby([district_column, progress_code_column]).size().unstack(fill_value=0)
progress_summary.columns = progress_summary.columns.map(progress_code_mapping)

# Analysis of Main Progress Code by Landowner Type with mapped values
landowner_progress_summary = data.groupby(['Landowner_Type', progress_code_column]).size().unstack(fill_value=0)
landowner_progress_summary.columns = landowner_progress_summary.columns.map(progress_code_mapping)

# Sorting all centers by Landowner, District, and Translated Progress Code
sorted_centers = data.sort_values(by=[landowner_column, district_column, 'Translated_Progress_Code', center_name_column])

# Select relevant columns for output
sorted_centers_output = sorted_centers[[landowner_column, district_column, 'Translated_Progress_Code', center_name_column, 'Has_Doc_A', 'Has_Doc_B', 'Has_Doc_C']]

# Calculate districts for each type of landowner
landowner_districts = data.groupby(landowner_column)[district_column].unique()
landowner_centers = data.groupby(landowner_column)[center_name_column].count()  # Number of centers

# Calculate the percentage of each type of landowner
landowner_counts = data[landowner_column].value_counts()
landowner_percentages = (landowner_counts / landowner_counts.sum()) * 100

# Calculate the total number of districts that a landowner spans across
total_districts_per_landowner = data.groupby(landowner_column)[district_column].apply(lambda x: x.nunique())


# Prepare data for districts in a more readable format
landowner_districts_df = pd.DataFrame({
    'Landowner': landowner_districts.index,
    'Districts': landowner_districts.apply(lambda x: ', '.join(x)),

})

# Prepare the DataFrame for output
landowner_summary_df = pd.DataFrame({
    'Landowner': landowner_districts.index,
    'Total_Centers': landowner_centers.values,
    'Total_Districts': total_districts_per_landowner.values,  # Total number of districts spanned by landowner
    'Districts_Spanned': landowner_districts.values,  # Number of unique districts spanned
})


# Save all analyses to the same Excel file with different sheets
output_file_path = 'Doc_Statement_And_Progress_Code_Analysis.xlsx'  # replace with your desired save path

with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    # Write existing summaries
    doc_summary.to_excel(writer, sheet_name='Doc_Statement_Analysis', index=False)
    progress_summary.to_excel(writer, sheet_name='Progress_Code_Analysis')
    landowner_progress_summary.to_excel(writer, sheet_name='Landowner_Progress_Analysis')
    landowner_doc_summary.to_excel(writer, sheet_name='Landowner_Doc_Analysis', index=False)
    sorted_centers_output.to_excel(writer, sheet_name='Sorted_Centers', index=False)

    # Write landowner percentages to a new sheet
    landowner_percentages.to_frame(name='Percentage').to_excel(writer, sheet_name='Landowner_Percentages')

    # Write landowner districts to another new sheet
    landowner_districts_df.to_excel(writer, sheet_name='Landowner_Districts', index=False)
    landowner_summary_df.to_excel(writer, sheet_name='Landowner_Percentages', index=False)


# Display the results (optional)
print(doc_summary)
print(progress_summary)
print(landowner_progress_summary)
print(landowner_doc_summary)
print(sorted_centers_output)



# Print the results to the terminal
print("Percentage of each type of landowner:")
print(landowner_percentages)
print("\nDistricts covered by each type of landowner:")
for landowner, districts in landowner_districts.items():
    print(f"{landowner}: {', '.join(districts)}")
