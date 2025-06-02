from Pre_Analysis import load_dataframes
import pandas as pd
from pathlib import Path

#Load dataframes
dataframes = load_dataframes()

#Output directory (same as in Pre_Analysis for saving results but can be changed as needed.)
#Todo:Needs to be updated for each project.
output_dir = Path(r"F:\PROJECTS\1765 Forward Capacity Market (FCM) 2024\EVT 2024 Projects\Riazul\Hurtubise&Sons_SiteID541065_LocID566381\Cx Work\Analysis\Python")

#Holiday dates to exclude
#Todo:Update the list dates below based on your metering period
exclude_dates = pd.to_datetime([
    "2025-04-18",  # Friday
    "2025-04-21",  # Monday
    "2025-05-26"   # Monday
])

#Removing holidays from each DataFrame in the dictionary
for name, df in dataframes.items():
    df = df[~df['DateTime'].dt.normalize().isin(exclude_dates)]
    dataframes[name] = df

#CF Calculation
cf_results = []  #Empty list to store results in

for sheet_name, df in dataframes.items():
    try:
        #Prompt for threshold input
        threshold = float(input(f"Enter Lumen Intensity Value to Calculate CF_{sheet_name}: ")) #This is the threshold above which we know the fixture is on. This value is estimated from the intensity plots generated in the Pre_Analysis code and is used to account for any daylighting that might affect meter data. Use 0 if no daylighting is interfering in the space.

        #Overall CF
        total_count = len(df)
        above_threshold_count = (df['Intensity_lum/ft²'] > threshold).sum()
        overall_cf = round(above_threshold_count / total_count, 3) if total_count > 0 else None #Will give N/A value if it can't find any data points above threshold. In which case double check threshold.

        #Summer CF (weekday, 1PM–5PM) #If the metering period is large enough to include month restriction please add below. But it won't be necessary in most projects because the metering period is usually 2 weeks.
        summer_peak_mask = (
            df['DateTime'].dt.weekday < 5 &  # Weekday
            df['DateTime'].dt.hour.between(13, 16)  # 1PM to 4:59PM
        )
        summer_data = df[summer_peak_mask]
        summer_cf = None #Initializing the variable
        if not summer_data.empty:
            summer_cf = round((summer_data['Intensity_lum/ft²'] > threshold).sum() / len(summer_data), 3)

        #Winter CF (weekday, 5PM–7PM)
        winter_peak_mask = (
            df['DateTime'].dt.weekday < 5 &  # Weekday
            df['DateTime'].dt.hour.between(17, 18)  # 5PM to 6:59PM
        )
        winter_data = df[winter_peak_mask]
        winter_cf = None #Initializing the variable
        if not winter_data.empty:
            winter_cf = round((winter_data['Intensity_lum/ft²'] > threshold).sum() / len(winter_data), 3)

        #Storing results
        cf_results.append({
            'Sheet': sheet_name,
            'Threshold': threshold,
            'Overall CF': overall_cf if overall_cf is not None else 'N/A',
            'Summer CF': summer_cf if summer_cf is not None else 'N/A',
            'Winter CF': winter_cf if winter_cf is not None else 'N/A'
        })

    except ValueError:
        print(f"Invalid input for sheet {sheet_name}. Skipping.")

cf_df = pd.DataFrame(cf_results)
cf_df.to_csv(output_dir / "Coincidence_Factor_Summary.csv", index=False) #Todo:File name for output can be changed here
