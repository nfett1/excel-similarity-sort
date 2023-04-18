Project Title

Duplicate String Finder
Description

This script is designed to find duplicates of strings in an Excel file. It prompts the user to enter the file path of the Excel file and the name of the column to find similarities in. The user can also enter a similarity threshold and a scorer to use. The script then generates a new Excel file containing the results.
Dependencies

    pandas
    numpy
    rapidfuzz

Installation

    Clone the repository to your local machine.

    Install the dependencies using pip:

    pip install -r requirements.txt

Usage

To run the script, navigate to the directory where the script is located and run the following command:

python duplicate_string_finder.py

Follow the prompts to enter the file path of the Excel file, the name of the column to find similarities in, the similarity threshold, and the scorer to use.

The script will generate a new Excel file containing the results in the same directory as the input file.
Examples

lua

Enter the file path: /path/to/file.xlsx
Enter the column name to find similaries in: Name
Enter the similarity threshold to apply ( 0 - 100 ): 80
Enter the scorer to use ( token_set_ratio, token_sort_ratio, partial_ratio, partial_token_set_ratio, partial_token_sort_ratio, WRatio, QRatio ): token_set_ratio
Successfully created file /path/to/file_output.xlsx

Credits

This project was created by [Nati Solomon Fett] for educational purposes.