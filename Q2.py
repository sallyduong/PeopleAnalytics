# %%
import pandas as pd
from scipy.stats import chi2_contingency

# %%
# Import data
df = pd.read_excel("Dataset_2.xlsx")

# Add a new column indicating the year the application was submitted
df['App_Year'] = df['Date of Application'].dt.year

# %%
# Define a mapping for the stages
stage_mapping = {
    'New Application': 1,
    'Phone Screen': 2,
    'In-House Interview': 3,
    'Offer Sent': 4,
    'Offer Accepted': 5
}

# Map the 'Furthest Recruiting Stage Reached' column to its numerical value
df['Stage_Num'] = df['Furthest Recruiting Stage Reached'].map(stage_mapping)

# Identify candidates (i) whose Furthest Recruiting Stage Reached was In-House Interview or beyond
df_filtered = df[df['Stage_Num'] >= 3]

# Identify candidates (ii) who applied from the Application Sources of Career Fair or Campus Event only
df_filtered  = df_filtered [(df_filtered['Application Source'] == 'Career Fair') | (df_filtered['Application Source'] == 'Campus Event')]

# %%
# Calculate total number of applicants (assuming each row is an applicant) per year
total_apps = df.groupby('App_Year').size().reset_index(name='Count_Total')

# Calculate the total numer of applicants that reached In-House Interview or beyond per year
reached_apps = df_filtered.groupby(['App_Year', 'Application Source']).size().reset_index(name='Count_Reached')

# Join the two dataframes
df_joined = pd.merge(total_apps, reached_apps, on='App_Year', how='inner')

# Calculate 'Reached In-House Interview'
df_joined['Reached_In_House_Interview'] = (df_joined['Count_Reached'] / df_joined['Count_Total']) * 100

# Pivot the table to form the contingency table
contingency_table = df_joined.pivot(index='App_Year', columns='Application Source', values='Reached_In_House_Interview')

# %%
# Perform chi-squared tests year-over-year
years = contingency_table.index
for i in range(len(years) - 1):
    year1 = years[i]
    year2 = years[i + 1]

    contingency_year1 = contingency_table.loc[year1]
    contingency_year2 = contingency_table.loc[year2]

    chi2_stat, p_val, dof, expected = chi2_contingency([contingency_year1, contingency_year2])

    print(f"\nChi-squared test between {year1} and {year2}:")
    print(f"Chi2 Statistic: {chi2_stat}")
    print(f"P-value: {p_val}")
    print(f"Degrees of Freedom: {dof}")
    print("Expected Frequencies:")
    print(expected)

# %% [markdown]
# ## Summary of the results
# There is no significant difference in the distribution of Application Sources (Career Fair, Campus Event) between the specified pairs of years (2016 vs. 2017 and 2017 vs. 2018)
#
# The p-value of 1.0 suggests that there is not enough evidence to reject the null hypothesis. In both tests, the Application Sources for the given years is consistent with what would be expected by chance.
