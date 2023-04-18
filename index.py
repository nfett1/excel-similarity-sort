import os
import pandas as pd
import numpy as np
from rapidfuzz import process,fuzz

def loadFile( fileName ):
    # Load Excel file into a DataFrame
    return pd.read_excel( fileName, sheet_name='Sheet1' )

def getColumn( df, column_name):
    # Get the column of interest from the input file
    return df.loc[:, column_name].tolist()

def get_similarities(elements, scorer, score_cutoff):
    # Create a list of lists to store the results
    results = [[name, [], 0] for name in elements]
    # Get the similar strings
    similarity_check = process.cdist(elements,elements, scorer=getattr(fuzz, scorer), score_cutoff=float(score_cutoff), workers=-1)

    duplicates_list = []
    for distances in similarity_check:
        # Get indices of duplicates
        indices = np.argwhere(~np.isin(distances, [100, 0])).flatten()
        # Get names from indices
        names = list(map(elements.__getitem__, indices))
        duplicates_list.append(names)

    # Create dataframe using the data generated above
    df = pd.DataFrame({'name': elements, 'similarities': duplicates_list})
    df['similarity_count'] = df.similarities.str.len()
    return df

def createOutputFile( df, file_path ):
    # Create a new Excel file
    output_file_path = os.path.splitext(file_path)[0] + '_output.xlsx'
    df.to_excel(output_file_path, index=False)
    print(f"Successfully created file {output_file_path}")

def main():
    file_path = input("Enter the file path: ")
    try:
        df = loadFile(file_path)
        print("columns found: ", df.columns)
        column_number = input("Enter the column name to find similaries in: ")
        element = getColumn(df, column_number)
        score_threshold = input("Enter the similarity threshold to apply ( 0 - 100 ): ")
        scorer = input("Enter the scorer to use ( token_set_ratio, token_sort_ratio, partial_ratio, partial_token_set_ratio, partial_token_sort_ratio, WRatio, QRatio ): ")
        df = get_similarities(element, scorer, score_threshold)
        createOutputFile(df, file_path)
        
    except FileNotFoundError:
        print(f"Error: File {file_path} not found.")

main()