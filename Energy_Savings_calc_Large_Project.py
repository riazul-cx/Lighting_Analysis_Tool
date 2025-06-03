import pandas as pd
from pathlib import Path

#Get CF data from earlier analysis
#Todo: Update file paths for each project. Comment it out if CF is coming from an excel input below.
CF_input_file_path = r"F:\PROJECTS\1765 Forward Capacity Market (FCM) 2024\EVT 2024 Projects\Riazul\Hurtubise&Sons_SiteID541065_LocID566381\Cx Work\Analysis\Python\Coincidence_Factor_Summary.csv"
CF_df = pd.read_csv(CF_input_file_path)

#Average CF calc based on space type. This is if multiple meters have been installed in similar space tyeps. Comment (Shortcut: Ctrl+/) it out if not necessary.
#Todo:The below chunk only work for space types with a naming scheme like "H69-Calving Barn". For different naming schemes it will need some tweaks.

CF_df['Space Type'] = CF_df['Sheet'].str.split('-').str[-1].str.strip()
Cf_summary_by_space = CF_df.groupby('Space Type')[['Overall CF', 'Summer CF', 'Winter CF']].mean().reset_index()
Cf_summary_by_space = Cf_summary_by_space.round(3)
print(Cf_summary_by_space)

#Lets get some of the other factors from an Excel file
excel_path = input("Enter full Excel file path (including .xlsx and no quotation marks): ").strip()
sheet_name = input("Enter sheet name to use: ").strip()

df = pd.read_excel(excel_path, sheet_name=sheet_name)

print("\nAvailable columns in sheet:")
print(df.columns.tolist())

space_col = input("Enter column name for Space Type: ").strip()
baseline_col = input("Enter column name for Baseline Fixtures No.: ").strip()
efficient_col = input("Enter column name for Efficient Fixtures No.: ").strip()
efficient_watt_col = input("Enter column name for Efficient Wattage: ").strip()
baseline_watt_col = input("Enter column name for Baseline Wattage: ").strip()
#baseline_hours_col = input("Enter column name for Baseline Annual Hours of Operation: ").strip() #Todo: If baseline hours of op are different from efficient case
# Overall_cf_col = input("Enter column name for Overall CF (if applicable): ").strip()  #Optional CF from Excel – uncomment if needed
# Summer_cf_col = input("Enter column name for Summer CF (if applicable): ").strip()  #Optional CF from Excel – uncomment if needed
# Winter_cf_col = input("Enter column name for Winter CF (if applicable): ").strip()  #Optional CF from Excel – uncomment if needed

#Creating lighting_df
lighting_df = df[[space_col, baseline_col, efficient_col, efficient_watt_col, baseline_watt_col]].copy()
lighting_df.columns = ["Space Type", "Baseline Fixtures", "Efficient Fixtures", "Efficient Wattage", "Baseline Wattage"]
# lighting_df['Baseline Annual Hours'] = df[baseline_hours_col]  #Todo:Uncomment if baseline hours of op are different
# lighting_df['Overall CF'] = df[Overall_cf_col]  #Todo:Uncomment if using CF directly from Excel
# lighting_df['Summer CF'] = df[Summer_cf_col]  #Todo:Uncomment if using CF directly from Excel
# lighting_df['Winter CF'] = df[Winter_cf_col]  #Todo:Uncomment if using CF directly from Excel

#Merging CFs if they are coming from meter data #Todo: Comment it out if CF is coming from an excel input below.
lighting_df = lighting_df.merge(Cf_summary_by_space, how='left', on='Space Type')

#Prompt for CFs if any are missing #Todo: Best case is check your excel file to make sure no CFs are missing
for i, row in lighting_df[lighting_df['Overall CF'].isna()].iterrows():
    space = row['Space Type']
    print(f"\nNo CF data found for space type: '{space}'")
    try:
        overall_cf = float(input(f"Enter Overall CF for '{space}': "))
        summer_cf = float(input(f"Enter Summer CF for '{space}': "))
        winter_cf = float(input(f"Enter Winter CF for '{space}': "))
    except ValueError:
        print("Invalid input. Using 1.0 for all CF values.")
        overall_cf = summer_cf = winter_cf = 1.0

    lighting_df.at[i, 'Overall CF'] = round(overall_cf, 3)
    lighting_df.at[i, 'Summer CF'] = round(summer_cf, 3)
    lighting_df.at[i, 'Winter CF'] = round(winter_cf, 3)

#TRM Parameters
WHF_d = float(input("\nEnter waste heat factor for demand to account for cooling savings from efficient lighting (WHF_d): "))
WHF_e = float(input("Enter waste heat factor for energy to account for cooling savings from efficient lighting (WHF_e): "))
OA = float(input("Enter the average percent of supply air that is Outside Air (OA) (not in %): "))
AR = float(input("Enter aspect ratio factor factor from TRM (AR): "))
HF = float(input("Enter ASHRAE heating factor (HF) from TRM: "))
DFH = float(input("Enter percent of lighting in heated spaces either from TRM or if known from site visit (DFH): "))
Efficiency_heat = float(input("Enter heating system efficiency (not in %): "))

#Now we have all the variables to do out savings analysis finally!
lighting_df['Annual Hours'] = lighting_df['Overall CF'] * 8760

#Todo: Below kWh savings equation is if both baseline and efficient case have same hours of operation
lighting_df['Energy Savings (kWh)'] = round(((lighting_df['Baseline Fixtures'] * lighting_df['Baseline Wattage']) -(lighting_df['Efficient Fixtures'] * lighting_df['Efficient Wattage']))/1000*
                                            lighting_df['Annual Hours'] * WHF_e, 3)

#Todo: Below kWh savings equation is if baseline and efficient case have different hours of operation
#lighting_df['Energy Savings (kWh)'] = round(((lighting_df['Baseline Fixtures'] * lighting_df['Baseline Wattage'] * lighting_df['Baseline Annual Hours']) -(lighting_df['Efficient Fixtures'] * lighting_df['Efficient Wattage'] * lighting_df['Annual Hours']))/1000 * WHF_e, 3)

lighting_df['Summer Demand Savings (kW)'] = round(((lighting_df['Baseline Fixtures'] * lighting_df['Baseline Wattage']) - (lighting_df['Efficient Fixtures'] * lighting_df['Efficient Wattage']))/1000*
    lighting_df['Summer CF'] * WHF_d, 3)

lighting_df['Winter Demand Savings (kW)'] = round(((lighting_df['Baseline Fixtures'] * lighting_df['Baseline Wattage']) - (lighting_df['Efficient Fixtures'] * lighting_df['Efficient Wattage'])) / 1000 *
    lighting_df['Winter CF'] * WHF_d, 3)

lighting_df['Natural Gas Savings (MMBtu)'] = round((lighting_df['Energy Savings (kWh)'] / WHF_e) * 0.003412 * (1 - OA) * AR * HF * DFH / Efficiency_heat, 3)

#Save Results of Analysis

output_dir_input = input("Enter the folder path where the output CSV should be saved: ").strip()
output_dir = Path(output_dir_input)
output_dir.mkdir(parents=True, exist_ok=True) #Ensure the directory exists
lighting_df.to_csv(output_dir / "Lighting Analysis.csv", index=False)
print("\nLighting Analysis saved to: {output_dir / 'Lighting Analysis.csv'}.")
