import pandas as pd
import re
from collections import Counter

# def readexcel():
#     filepath = "C:\Users\Dell\Downloads\coding challenge test.xlsx"
#
#     dataframe = pd.read_excel(filepath, engine='openpyxl')
#
#     comments = dataframe['Additional comments']
#
#     group_counter = Counter()
#
#     pattern

import pandas as pd

# print(np.__version__)
import re
from collections import Counter

def extract_groups(comments):
    try:
        pattern = re.compile(r"Groups : \[code\]<I>(.*?)<\/I>\[\/code\]")
        groups = []
        for comment in comments:
            if pd.notna(comment):
                matches = pattern.findall(comment)
                for match in matches:
                    groups.extend([group.strip() for group in match.split(',')])
        return groups
    except Exception as e:
        print(e)

def count_group_occurrences(input_file, sheet_name):

    try:
        # Read the Excel file
        df = pd.read_excel(input_file, sheet_name=sheet_name,engine='openpyxl')

        # Extract comments
        comments = df['Additional comments'].dropna()

        # Extract groups
        groups = extract_groups(comments)

        # Count occurrences
        group_counter = Counter(groups)

        # Create the output DataFrame
        output_df = pd.DataFrame(group_counter.items(), columns=['Group name', 'Number of occurrences'])
        output_df = output_df.sort_values(by='Number of occurrences', ascending=False)

        return output_df
    except Exception as e:
        print(e)

# Input file
input_file = "C:\\Users\\Dell\\Downloads\\coding challenge test.xlsx"
sheet_name = 'Input Data sheet'

output_df = count_group_occurrences(input_file, sheet_name)
print(output_df)
output_filepath = "Group_occurrences.xlsx"
output_df.to_excel(output_filepath,index=False)