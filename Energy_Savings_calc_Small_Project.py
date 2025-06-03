import pandas as pd
from pathlib import Path

#Get CF data from earlier analysis
#Todo: Update file paths for each project
CF_input_file_path = r"F:\PROJECTS\1765 Forward Capacity Market (FCM) 2024\EVT 2024 Projects\Riazul\Hurtubise&Sons_SiteID541065_LocID566381\Cx Work\Analysis\Python\Coincidence_Factor_Summary.csv"
CF_df = pd.read_csv(CF_input_file_path)

#Average CF calc based on space type. This is if multiple meters have been installed in similar space tyeps. Comment (Shortcut: Ctrl+/) it out if not necessary.
#Todo:The below chunk only work for space types with a naming scheme like "H69-Calving Barn". For different naming schemes it will need some tweaks. If stuck contact Riaz

CF_df['Space Type'] = CF_df['Sheet'].str.split('-').str[-1].str.strip()
Cf_summary_by_space = CF_df.groupby('Space Type')[['Overall CF', 'Summer CF', 'Winter CF']].mean().reset_index() #Todo:Taking the mean CF for similar space type. Comment this out if not necessary.
Cf_summary_by_space = Cf_summary_by_space.round(3)
print(Cf_summary_by_space) #Todo: Please to use the same space type names when prompted to enter later in the code.

#Optional: save to a new CSV
#cf_summary_by_space.to_csv( file path, index=False)

#Now to do some savings calculation
#Code for Small Projects

#Initialize lists for inputs
space_names = []
baseline_counts = []
efficient_counts = []
efficient_wattage = []
baseline_wattage =[]
#baseline_hours =[] #Todo: Uncomment this out if baseline hours are different from efficient case
#overall_CF = [] #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data. If all CFs are coming from TRM, comment out the code from CF_input_file_path to print(Cf_summary_by_space)
#summer_CF = [] #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.
#winter_CF = [] #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.

#Ask how many spaces the user wants to enter
num_spaces = int(input("\nEnter the number of space types to include: "))

for i in range(num_spaces):
    space = input(f"\nEnter name of space type #{i + 1}: ")
    baseline_count = float(input(f"Enter number of fixtures (Baseline) for '{space}': "))
    efficient_count = float(input(f"Enter number of fixtures (Efficient Case) for '{space}': "))
    efficient_watt = float(input(f"Enter wattage of fixtures (Efficient Case) for '{space}': "))
    baseline_watt = float(input(f"Enter wattage of fixtures (Baseline Case) for '{space}': "))
    #base_hours = float(input(f"Enter baseline hours of operation (Baseline Case) for '{space}': ")) #Todo: Uncomment this out if baseline hours are different from efficient case
    #over_CF_input = float(input(f"Enter Overall CF from TRM/Loadshape for '{space}': ")) #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.
    #sum_CF_input = float(input(f"Enter Summer CF from TRM/Loadshape for '{space}': ")) #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.
    #win_CF_input = float(input(f"Enter Winter CF from TRM/Loadshape for '{space}': ")) #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.

    #Appending input to lists
    space_names.append(space)
    baseline_counts.append(baseline_count)
    efficient_counts.append(efficient_count)
    efficient_wattage.append(efficient_watt)
    baseline_wattage.append(baseline_watt)
    #baseline_hours.append(base_hours) #Todo: Uncomment this out if baseline hours are different from efficient case
    #overall_CF.append(over_CF_input) #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.
    #summer_CF.append(sum_CF_input) #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.
    #winter_CF.append(win_CF_input) #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data.

#Creating DataFrame
lighting_df = pd.DataFrame({
    "Space Type": space_names,
    "Baseline Fixtures": baseline_counts,
    "Efficient Fixtures": efficient_counts,
    "Efficient Wattage": efficient_wattage,
    "Baseline Wattage": baseline_wattage #,
    #"Baseline Annual Hours": baseline_hours, #Todo: Uncomment this out if baseline hours are different from efficient case
    #"Overall CF": overall_CF, #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data. Uncomment the comma in the previous line too.
    #"Summer CF": summer_CF, #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data. Uncomment the comma in the previous line too.
    #"Winter CF": overall_CF #Todo: Uncomment this if the CFs are coming from a study or TRM instead of meter data. Uncomment the comma in the previous line too.
})

#Merge CF values into lighting_df #Todo: Comment it out if CF is coming from an input.
lighting_df = lighting_df.merge(Cf_summary_by_space,how='left',on='Space Type')

#Check for unmatched space types. #Todo: This is if some areas have CF from metering and for other space types we are using TRM CF. This is different from the commented out CFs where all the CFs are coming from TRM. Though they kind of do the samet thing.
for i, row in lighting_df[lighting_df['Overall CF'].isna()].iterrows():
    space = row['Space Type']
    print(f"\nNo CF data found for space type: '{space}'")
    try:
        overall_cf = float(input(f"Enter Overall CF for '{space}': "))
        summer_cf = float(input(f"Enter Summer CF for '{space}': "))
        winter_cf = float(input(f"Enter Winter CF for '{space}': "))
    except ValueError:
        print("Invalid input. Using 1 for all CF values.")
        overall_cf = summer_cf = winter_cf = 1.0

    # Assign the user-provided values to the DataFrame
    lighting_df.at[i, 'Overall CF'] = round(overall_cf, 3)
    lighting_df.at[i, 'Summer CF'] = round(summer_cf, 3)
    lighting_df.at[i, 'Winter CF'] = round(winter_cf, 3)

#TRM Parameters
WHF_d = float(input(f"\nEnter waste heat factor for demand to account for cooling savings from efficient lighting (WHF_d): "))
WHF_e = float(input(f"Enter waste heat factor for energy to account for cooling savings from efficient lighting (WHF_e): "))
OA = float(input(f"Enter the average percent of the supply air that is Outside Air (OA) from TRM (not in %): "))
AR = float(input(f"Enter the aspect ratio (AR) factor from TRM (not in %): "))
HF = float(input(f"Enter the ASHRAE heating factor (HF) from TRM: "))
DFH = float(input(f"Enter the percent of lighting in heated spaces either from TRM or if known from site visit (DFH) (not in %): "))
Efficiency_heat = float(input(f"Enter the heating system efficiency (not in %): "))

#Now we have all the variables to do out savings analysis finally!

lighting_df['Annual Hours'] = lighting_df['Overall CF']*8760

lighting_df['Energy Savings (kWh)'] = round(((lighting_df['Baseline Fixtures']*lighting_df['Baseline Wattage'])-(lighting_df['Efficient Fixtures']*lighting_df['Efficient Wattage']))/1000*lighting_df['Annual Hours']*WHF_e, 3)

#Todo: Below kWh savings equation is if baseline and efficient case have different hours of operation. Uncomment this out if needed and comment out the above line.
#lighting_df['Energy Savings (kWh)'] = round(((lighting_df['Baseline Fixtures'] * lighting_df['Baseline Wattage'] * lighting_df['Baseline Annual Hours']) -(lighting_df['Efficient Fixtures'] * lighting_df['Efficient Wattage'] * lighting_df['Annual Hours']))/1000 * WHF_e, 3)

lighting_df['Summer Demand Savings (kW)'] = round(((lighting_df['Baseline Fixtures']*lighting_df['Baseline Wattage'])-(lighting_df['Efficient Fixtures']*lighting_df['Efficient Wattage']))/1000*lighting_df['Summer CF']*WHF_d, 3)
lighting_df['Winter Demand Savings (kW)'] = round(((lighting_df['Baseline Fixtures']*lighting_df['Baseline Wattage'])-(lighting_df['Efficient Fixtures']*lighting_df['Efficient Wattage']))/1000*lighting_df['Winter CF']*WHF_d, 3)
lighting_df['Natural Gas Savings (MMBtu)'] = round((lighting_df['Energy Savings (kWh)']/WHF_e)*0.003412*(1-OA)*AR*HF*DFH/Efficiency_heat, 3)

#Save Results of Analysis

output_dir_input = input("Enter the folder path where the output CSV should be saved: ").strip()
output_dir = Path(output_dir_input)
output_dir.mkdir(parents=True, exist_ok=True) #Ensure the directory exists
lighting_df.to_csv(output_dir / "Lighting Analysis.csv", index=False)
print("\nLighting Analysis saved to: {output_dir / 'Lighting Analysis.csv'}.")

