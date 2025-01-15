import pandas as pd
import numpy as np
from datetime import timedelta

def process_sheet(input_file, input_sheet, output_sheet):
    # Read the raw data from the Excel file
    data = pd.read_excel(input_file, sheet_name=input_sheet)

    # Ensure required columns exist
    required_columns = ['StudySubjectID', 'Eye', 'ScanStartTime', 'DN_MSI', 'UpdateLongiPositions', 'EligibleQuant', 'Isincluded']
    if not all(col in data.columns for col in required_columns):
        raise ValueError(f"The required columns {required_columns} are missing in the data.")

    # Filter data to keep rows where 'PatientID' begins with 'AO'
    data = data[data['StudySubjectID'].str.startswith('AO')]

    # Create a combined identifier for patient_eye
    data['Patient_Eye'] = data['StudySubjectID'].astype(str) + '_' + data['Eye'].astype(str)

    # Convert ScanStartTime to datetime
    data['ScanStartTime'] = pd.to_datetime(data['ScanStartTime'])

    # Initialize a list to hold summary data
    summary_data = []

    # Define the minimum scans required per device
    Min_scans_for_device = 10  # Example value, set this as needed

    # Group by patient_eye and process each group
    for patient_eye, group in data.groupby('Patient_Eye'):
        # Extract 'DeviceID' from 'UniqueIdentifier' (substring before the 1st '_')
        group['DeviceID'] = group['UniqueIdentifier'].str.split('_').str[0]

        # Count rows for each 'DeviceID' in the group
        device_counts = group['DeviceID'].value_counts()

        # Filter the group to keep rows where 'DeviceID' count exceeds 'Min_scans_for_device'
        group = group[group['DeviceID'].isin(device_counts[device_counts > Min_scans_for_device].index)]

        if not group.empty:

            # Calculate 'is_study_eye' as the mean of 'Isincluded' rounded to the closest integer
            is_study_eye = round(group['Isincluded'].mean())

            # Sort by ScanStartTime
            group = group.sort_values(by='ScanStartTime')

            # Compute total days and number of tests
            total_days = (group['ScanStartTime'].max() - group['ScanStartTime'].min()).days + 1
            number_of_tests = group.shape[0]

            # Compute testing rate
            testing_rate = round(number_of_tests / total_days, 2) if total_days > 0 else 0

            # Compute the start and end of each calendar week
            first_date = group['ScanStartTime'].min()
            last_date = group['ScanStartTime'].max()

            # Compute total calendar weeks (include partial weeks)
            total_calendar_weeks = len(pd.date_range(start=first_date, end=last_date, freq='W-SUN'))

            # Compute Adherence Rate (Scans/Week
            Adherence_Rate = round(number_of_tests / total_calendar_weeks, 2)

            # Image quality metrics (DN_MSI)
            mean_msi = round(group['DN_MSI'].mean(), 2)
            # sd_msi = round(group['DN_MSI'].std(), 2)
            # median_msi = round(group['DN_MSI'].median(), 2)
            # iqr_msi = round(group['DN_MSI'].quantile(0.75) - group['DN_MSI'].quantile(0.25), 2)

            # Append summary
            summary_data.append({
                'Patient_Eye': patient_eye,
                'Study_Eye': is_study_eye,
                'Adherence_Rate': Adherence_Rate,
                'Mean DN_MSI': mean_msi,
                # 'SD DN_MSI': sd_msi,
                # 'Median DN_MSI': median_msi,
                # 'IQR DN_MSI': iqr_msi
            })

    # Convert summary data to DataFrame
    summary_df = pd.DataFrame(summary_data)

    study_eye_summary_df = summary_df[summary_df.Study_Eye == 1]
    not_study_eye_summary_df = summary_df[summary_df.Study_Eye == 0]

    # Write the summary to a new sheet in the Excel file
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        summary_df.to_excel(writer, sheet_name=output_sheet, index=False)

    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        study_eye_summary_df.to_excel(writer, sheet_name=output_sheet + "_study_eyes", index=False)

    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        not_study_eye_summary_df.to_excel(writer, sheet_name=output_sheet + "_not_study_eyes", index=False)

    print(f"Summary sheet {output_sheet} created successfully.")

# File and sheet information
file_path = r"\\172.17.102.175\Algorithm\Ilan\AO\Report3_January_2025\AO_from_DB_15012025.xlsx"

# Process 3rd sheet
process_sheet(file_path, 'AO_RawData', 'AO_Summary')