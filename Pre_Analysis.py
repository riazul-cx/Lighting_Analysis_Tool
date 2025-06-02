#This script is meant to calculate savings from efficient lighting FCM projects when meter data is available.
#Todo below are what needs to be modified when this code is being used for a new project. Please read through them craefully.

import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

#Todo:Ensure that all meter data are in one excel file. The column headers and sheet names are important. Ensure the column headers are in the general format: Index, DateTime, Intensity_lum/ft². Sheet name example: H69-Calving Barn; where H69 is the meter tag and Calving Barn is the space type.
#Todo: Sheet name, naming structure becomes important in the "Energy_savings" script
#Todo:If your headers are different some changes will be needed to the plot headers below. Check out the input file path below if you need an example of what the raw data file should look like.

#Todo: Update below file paths for each project
input_file_path = Path(r"F:\PROJECTS\1765 Forward Capacity Market (FCM) 2024\EVT 2024 Projects\Riazul\Hurtubise&Sons_SiteID541065_LocID566381\Cx Work\Analysis\Combined Data_Python Import.xlsx")
output_dir = Path(r"F:\PROJECTS\1765 Forward Capacity Market (FCM) 2024\EVT 2024 Projects\Riazul\Hurtubise&Sons_SiteID541065_LocID566381\Cx Work\Analysis\Python")

#Create the dataframes, each sheet of the excel file gets converted into a dataframe
def load_dataframes():
    xls = pd.ExcelFile(input_file_path)
    sheet_names = xls.sheet_names
    dataframes = {}

    for sheet in sheet_names:
        df = pd.read_excel(input_file_path, sheet_name=sheet)
        df['DateTime'] = pd.to_datetime(df['DateTime'])
        dataframes[sheet] = df

    return dataframes

#The below code will run only if this script is run directly and not when the script is imported. This is impotant cause the time consuming plots don't need to be made everythime we import the dataframes into another script.
if __name__ == "__main__":
    output_dir.mkdir(parents=True, exist_ok=True) #Ensure output directory exists
    dataframes = load_dataframes() #Load data
    stats_list = [] #Initiating empty list to store values

    #Plot the data and generate basic stats
    for sheet, df in dataframes.items():
        min_date = df['DateTime'].min() - pd.Timedelta(days=1)
        max_date = df['DateTime'].max() + pd.Timedelta(days=1)

        plt.figure(figsize=(15, 6))
        plt.plot(df['DateTime'], df['Intensity_lum/ft²'], label='Intensity_lum/ft²', color='orange')
        plt.xlabel('DateTime')
        plt.ylabel('Intensity_lum/ft²')
        plt.title(f'Intensity Over Time - Sheet: {sheet}')
        plt.xlim(min_date, max_date)
        plt.grid(True)
        plt.tight_layout()
        plt.savefig(output_dir / f"{sheet}_Intensity_Plot.png")
        plt.close()

        # Stats
        stats_list.append({
            'Sheet': sheet,
            'First Sample Date': df["DateTime"].min(),
            'Last Sample Date': df["DateTime"].max(),
            'Max Intensity': df['Intensity_lum/ft²'].max(),
            'Min Intensity': df['Intensity_lum/ft²'].min(),
            'Mean Intensity': df['Intensity_lum/ft²'].mean(),
            'Standard Deviation': df['Intensity_lum/ft²'].std()
        })

    stats_df = pd.DataFrame(stats_list)
    stats_df.to_csv(output_dir / "Intensity_Statistics_Summary.csv", index=False) #File name for output can be changed here

    print(f"All plots and statistics saved to: {output_dir}")
