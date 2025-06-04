import numpy as np
import pandas as pd

CF_input_file_path = r"F:\PROJECTS\1765 Forward Capacity Market (FCM) 2024\EVT 2024 Projects\Riazul\Hurtubise&Sons_SiteID541065_LocID566381\Cx Work\Analysis\Python Analysis\Coincidence_Factor_Summary.csv"
cf_df = pd.read_csv(CF_input_file_path)

#Convert CF columns to numeric (in case any "N/A" strings exist)
cf_df['Overall CF'] = pd.to_numeric(cf_df['Overall CF'], errors='coerce')
cf_df['Summer CF'] = pd.to_numeric(cf_df['Summer CF'], errors='coerce')
cf_df['Winter CF'] = pd.to_numeric(cf_df['Winter CF'], errors='coerce')

#Standard deviations
overall_cf_std = cf_df['Overall CF'].std()
summer_cf_std = cf_df['Summer CF'].std()
winter_cf_std = cf_df['Winter CF'].std()

#Sample sizes
n_overall = cf_df['Overall CF'].count()
n_summer = cf_df['Summer CF'].count()
n_winter = cf_df['Winter CF'].count()

#Standard errors
overall_cf_se = overall_cf_std / np.sqrt(n_overall)
summer_cf_se = summer_cf_std / np.sqrt(n_summer)
winter_cf_se = winter_cf_std / np.sqrt(n_winter)

#Print results
print(f"\nStandard Deviation and Standard Error of CFs:")
print(f"Overall CF - SD: {overall_cf_std:.4f}, SE: {overall_cf_se:.4f}")
print(f"Summer CF  - SD: {summer_cf_std:.4f}, SE: {summer_cf_se:.4f}")
print(f"Winter CF  - SD: {winter_cf_std:.4f}, SE: {winter_cf_se:.4f}")
