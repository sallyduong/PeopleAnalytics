# %%
import pandas as pd
import matplotlib.pyplot as plt
from pandas.plotting import table

# %%
# Import data and fix headers
df_OfferResponse = pd.read_excel("Dataset.xlsx", sheet_name = 'Offer Response Data')

# Combine first two non-null rows of headers to differentiate education item
df_Activity = pd.read_excel("Dataset.xlsx", sheet_name = 'Recruiting Activity Data', skiprows=1, header = [0,1])
df_Activity.columns = df_Activity.columns.map('_'.join)

# Clean up column names
df_Activity.columns = df_Activity.columns.str.replace(r'.*level_0_', '', regex=True)

# %%
# Functions for converting degrees to rank order, and vice versa
def degree_rank(degree):
    """
    Description:
    Converts educational degrees to a rank for comparison.

    Parameters:
    - degree (string): type of education

    Returns:
    (int): rank of the education
    """
    if pd.isna(degree):
        return None
    elif degree == 'PhD':
        return 1
    elif (degree == 'Masters' or degree == 'JD'):
        return 2
    elif (degree == 'Bachelors'):
        return 3
    else:
        return -1

def convert_rank_to_degree(rank):
    """
    Description:
    Converts educational ranks to degrees. Combines Masters and JDs to a single enumeration.

    Parameters:
    - rank (int): rank of education

    Returns:
    (str): type of the education
    """
    if pd.isna(rank):
        return None
    elif rank == 1:
        return 'PhD'
    elif rank == 2:
        return 'Masters/JD'
    elif rank == 3:
        return 'Bachelors'
    else:
        return -1

# %%
# Convert degrees to rank order
df_Activity['Education 1_Degree_Rank'] = df_Activity['Education 1_Degree'].apply(lambda x: degree_rank(x))
df_Activity['Education 2_Degree_Rank'] = df_Activity['Education 2_Degree'].apply(lambda x: degree_rank(x))
df_Activity['Education 3_Degree_Rank'] = df_Activity['Education 3_Degree'].apply(lambda x: degree_rank(x))

# Compare up to three degrees per applicant and identify the highest level of education, with 1 being the highest
df_Activity['Highest_Degree'] = df_Activity[['Education 1_Degree_Rank', 'Education 2_Degree_Rank', 'Education 3_Degree_Rank']].min(axis=1)

# Convert the rank of education back to a degree name
df_Activity['Highest_Degree'] = df_Activity['Highest_Degree'].apply(lambda x: convert_rank_to_degree(x))

# %%
# Join the two datasets to get Offer Resoponse Data
df_Activity_Response = pd.merge(df_Activity, df_OfferResponse, on='Candidate ID Number', how='left')

# Update the 'Furthest Recruiting Stage Reached' if the offer was accepted
df_Activity_Response.loc[df_Activity_Response['Offer Decision'] == 'Offer Accepted', 'Furthest Recruiting Stage Reached'] = df_Activity_Response['Offer Decision']

# Clean enumeration in 'Furthest Recruiting Stage Reached'
df_Activity_Response['Furthest Recruiting Stage Reached'] = df_Activity_Response['Furthest Recruiting Stage Reached'].replace('In-house Interview', 'In-House Interview')

# %%
# Output dataset for Question 2
df_Activity_Response.to_excel('Dataset_2.xlsx', index=False)

# %%
# Grouping by 'Furthest Recruiting Stage Reached', 'Department', and 'Highest_Degree'
# Count the number of rows in each group to get the count for the latest stage reached
grouped_df = df_Activity_Response.groupby(['Furthest Recruiting Stage Reached', 'Department', 'Highest_Degree']).size().reset_index(name='Count')

# %%

# Define a mapping for the stages
# 1) New Application 2) Phone Screen 3) In-House Interview 4) Offer Sent 5) Offer Accepted
stage_mapping = {
    'New Application': 1,
    'Phone Screen': 2,
    'In-House Interview': 3,
    'Offer Sent': 4,
    'Offer Accepted': 5
}

# Backfill counts for earlier stages in the funnel
# Iterate through each row
for index, row in grouped_df.iterrows():
    current_stage = row['Furthest Recruiting Stage Reached']
    current_count = row['Count']

    # Update count for earlier stages
    for stage, value in stage_mapping.items():

        if value < stage_mapping[current_stage]:
            # Find the corresponding row for the earlier stage
            earlier_stage_row = grouped_df[
                (grouped_df['Department'] == row['Department']) &
                (grouped_df['Highest_Degree'] == row['Highest_Degree']) &
                (grouped_df['Furthest Recruiting Stage Reached'] == stage)
            ]

            # Update the count
            grouped_df.loc[earlier_stage_row.index, 'Count'] += current_count

# %%
# Re-order columns
grouped_df = grouped_df[['Department', 'Highest_Degree', 'Furthest Recruiting Stage Reached', 'Count']]

# Define the order for 'Furthest Recruiting Stage Reached' outlined previously in stage_mapping
custom_order = sorted(stage_mapping.keys(), key=lambda x: stage_mapping[x])

# Convert 'Furthest Recruiting Stage Reached' to a categorical type with the custom order
grouped_df['Furthest Recruiting Stage Reached'] = pd.Categorical(
    grouped_df['Furthest Recruiting Stage Reached'],
    categories=custom_order,
    ordered=True
)

# Sort the DataFrame by 'Department', 'Highest_Degree', and the custom-ordered 'Furthest Recruiting Stage Reached'
grouped_df = grouped_df.sort_values(by=['Department', 'Highest_Degree', 'Furthest Recruiting Stage Reached'])
# %%
# Calculate the conversion rate based on the grouped DataFrame using the value from the previous row
grouped_df['Conversion Rate (%)'] = (grouped_df['Count'] / grouped_df.groupby(['Department', 'Highest_Degree'])['Count'].shift(1)) * 100

# Format 'Conversion Rate (%)' column to have no decimal places and add '%' sign
grouped_df['Conversion Rate (%)'] = grouped_df['Conversion Rate (%)'].apply(lambda x: f'{x:.0f}%' if not pd.isna(x) else '')
# %%

# Create a plot with the table
fig, ax = plt.subplots(figsize=(10, 3))
ax.axis('off')
tab = table(ax, grouped_df, loc='center', colWidths=[0.2] * len(grouped_df.columns))
tab.auto_set_font_size(False)
tab.set_fontsize(8)

# Save the plot as a PDF file
output_pdf_path = 'Q1.pdf'
plt.tight_layout()  # Adjust layout for better spacing
plt.savefig(output_pdf_path, bbox_inches='tight', pad_inches=0.1)
plt.close()
