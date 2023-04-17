import pandas as pd
from rapidfuzz import fuzz

# Load Excel file into a DataFrame
df = pd.read_excel('book1.xlsx', sheet_name='Sheet1')

# Get the column of interest
column_name = 'Case description'
column = df.iloc[:, 0].tolist()

# Calculate the similarity between each pair of strings in the column
similarities = []
for i in range(len(column)):
    for j in range(i+1, len(column)):
        similarity = fuzz.token_sort_ratio(column[i], column[j])
        print( i, j, similarity)
        similarities.append((i, j, similarity))

# Create a dictionary to store the groups
groups = {}
for i, string in enumerate(column):
    group_id = None
    for j, k, similarity in similarities:
        print(i, j, k, similarity)
        if i in (j, k) and similarity >= 60:
            print("group_id : ", group_id)
            if group_id is None:
                print("here 2")
                group_id = j if i == k else k
                print(group_id)
            else:
                print("here 3")
                # if group_id not in groups:
                #     groups[group_id] = []
                groups[group_id].append(i) #TODO: this is not working
    if group_id is None:
        print("here 4")
        groups[i] = [i]

# Create a new DataFrame with the grouped rows
grouped_rows = []
for group_id, indices in groups.items():
    group_strings = [column[i] for i in indices]
    grouped_row = {
        column_name: '; '.join(group_strings),
        'Group ID': group_id,
        'Group Size': len(indices),
        'Similar case descriptions': [df.loc[i, 'SAP error opening form'] for i in indices]
    }
    grouped_rows.append(grouped_row)
grouped_df = pd.DataFrame(grouped_rows)

# Save the grouped DataFrame to a new Excel file
grouped_df.to_excel('example_grouped.xlsx', index=False)