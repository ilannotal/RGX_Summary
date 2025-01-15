import pandas as pd
import numpy as np
from datetime import timedelta

def process_sheet(input_file, input_sheet, output_sheet):
    # Read the raw data from the Excel file
    data = pd.read_excel(input_file, sheet_name=input_sheet)

    # Ensure required columns exist
    required_columns = ['StudySubjectID', 'Eye', 'ScanStartTime', 'DN_MSI']
    if not all(col in data.columns for col in required_columns):
        raise ValueError(f"The required columns {required_columns} are missing in the data.")

    # Create a combined identifier for patient_eye
    data['Patient_Eye'] = data['StudySubjectID'].astype(str) + '_' + data['Eye'].astype(str)

    # Convert ScanStartTime to datetime
    data['ScanStartTime'] = pd.to_datetime(data['ScanStartTime'])

    # Initialize a list to hold summary data
    summary_data = []

    # Group by patient_eye and process each group
    for patient_eye, group in data.groupby('Patient_Eye'):
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

        # Compute calendar weeks where at least one test occurred
        calendar_weeks = group['ScanStartTime'].dt.to_period('W-SUN').nunique()

        # Ensure weekly adherence rate does not exceed 100%
        weekly_adherence_rate = round(min(calendar_weeks / total_calendar_weeks * 100, 100),
                                      2) if total_calendar_weeks > 0 else 0

        # # Compute seven-day gaps without tests
        days_with_tests = set(group['ScanStartTime'].dt.date)
        # seven_day_periods = total_days // 7
        # gaps = 0
        #
        # for i in range(seven_day_periods):
        #     start_date = group['ScanStartTime'].min().date() + timedelta(days=i * 7)
        #     end_date = start_date + timedelta(days=6)
        #     if not any(start_date <= date <= end_date for date in days_with_tests):
        #         gaps += 1
        #
        # no_test_periods_ratio = round(gaps / seven_day_periods, 2) if seven_day_periods > 0 else 0

        # Compute all rolling 7-day periods
        start_date = group['ScanStartTime'].min().date()
        end_date = group['ScanStartTime'].max().date()

        # Iterate through each day in the range as the start of a 7-day window
        gaps = 0
        while start_date + timedelta(days=6) <= end_date:
            rolling_end_date = start_date + timedelta(days=6)
            if not any(start_date <= date <= rolling_end_date for date in days_with_tests):
                gaps += 1
            start_date += timedelta(days=1)  # Move to the next day

        # Compute the ratio of gaps
        rolling_periods = total_days - 6  # Total rolling 7-day periods
        no_test_periods_ratio = round(gaps / rolling_periods, 2) if rolling_periods > 0 else 0

        # Image quality metrics (DN_MSI)
        mean_msi = round(group['DN_MSI'].mean(), 2)
        sd_msi = round(group['DN_MSI'].std(), 2)
        median_msi = round(group['DN_MSI'].median(), 2)
        iqr_msi = round(group['DN_MSI'].quantile(0.75) - group['DN_MSI'].quantile(0.25), 2)

        # Append summary
        summary_data.append({
            'Patient_Eye': patient_eye,
            'Testing Rate': testing_rate,
            'Weekly Adherence Rate (%)': weekly_adherence_rate,
            'No-Test Periods Ratio': no_test_periods_ratio,
            'Mean DN_MSI': mean_msi,
            'SD DN_MSI': sd_msi,
            'Median DN_MSI': median_msi,
            'IQR DN_MSI': iqr_msi
        })

    # Convert summary data to DataFrame
    summary_df = pd.DataFrame(summary_data)

    # Write the summary to a new sheet in the Excel file
    with pd.ExcelWriter(input_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        summary_df.to_excel(writer, sheet_name=output_sheet, index=False)

    print(f"Summary sheet {output_sheet} created successfully.")

# File and sheet information
file_path = r"\\172.17.102.175\Algorithm\Ilan\RGX\RGX_summary_20250113.xlsx"

# # Process first sheet
# process_sheet(file_path, 'RGX-314-2103_RawData', 'RGX-314-2103_Summary')
#
# # Process second sheet
# process_sheet(file_path, 'RGX-314-5101_RawData', 'RGX-314-5101_Summary')

# Process 3rd sheet
process_sheet(file_path, 'AO_RawData', 'AO_Summary')