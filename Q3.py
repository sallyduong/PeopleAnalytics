# %%
import pandas as pd

# %%
# Stage Funnel Per Application Source
# Import data
df = pd.read_excel("Dataset_2.xlsx")

# %%
# Similar logic used in question 1

# Grouping by 'Furthest Recruiting Stage Reached', and 'Application Source'
# Counting the number of rows in each group to get the count for the latest stage reached
grouped_df = df.groupby(['Furthest Recruiting Stage Reached', 'Application Source']).size().reset_index(name='Count')

# Define a mapping for the stages
stage_mapping = {
    'New Application': 1,
    'Phone Screen': 2,
    'In-House Interview': 3,
    'Offer Sent': 4,
    'Offer Accepted': 5
}

# Iterate through each row
for index, row in grouped_df.iterrows():
    current_stage = row['Furthest Recruiting Stage Reached'] 
    current_count = row['Count']
    
    # Update count for earlier stages
    for stage, value in stage_mapping.items():

        if value < stage_mapping[current_stage]:
            # Find the corresponding row for the earlier stage
            earlier_stage_row = grouped_df[
                (grouped_df['Application Source'] == row['Application Source']) &
                (grouped_df['Furthest Recruiting Stage Reached'] == stage)
            ]
            
            # Update the count
            grouped_df.loc[earlier_stage_row.index, 'Count'] += current_count

# Re-order columns
grouped_df = grouped_df[['Application Source', 'Furthest Recruiting Stage Reached', 'Count']]

# Define the custom order for 'Furthest Recruiting Stage Reached'
custom_order = sorted(stage_mapping.keys(), key=lambda x: stage_mapping[x])

# Sort the DataFrame by 'Department' and then by 'Highest_Degree'
grouped_df = grouped_df.sort_values(by=['Application Source', 'Furthest Recruiting Stage Reached'])

# Convert the column to Categorical with the custom order
grouped_df['Furthest Recruiting Stage Reached'] = pd.Categorical(grouped_df['Furthest Recruiting Stage Reached'], categories=custom_order, ordered=True)

# Rename columns
grouped_df.rename(columns={'Furthest Recruiting Stage Reached': 'Stage'}, inplace=True)
grouped_df.rename(columns={'Count': 'Applicants'}, inplace=True)

# %%
# Output dataset for Question 3
grouped_df.to_excel('Dataset_3.xlsx', index=False)