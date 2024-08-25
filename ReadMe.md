## Project Title:
## Data Extraction and Transformation Using Python :)__
---
### Description: 
- This Python Script Filter Data From an Input Excel file based on specified row indices and save the Filtered Data to New Excel file

### Requirements:
+ Python 3.11.5
+ Pandas library ( install using pip install pandas )

### Usage:
 Place the input Excel file
+  (input_file333.xlsx) in the same directory as the script.

 Run the script
+ The filtered data will be saved as filtered_output.xlsx in the same directory.



### Table of contents


* [1. Script introduction](#Script-introduction)
* [2. Code Explanation](#Code-explanation)
    * [2.1 Importing Necessary Libraries](#Importing-Libraries)
    * [2.2 Input File and Required Columns](#Input-File-Required-Columns)
    * [2.3 Reading the Excel File](#reading-excel-file)
    * [2.4 Validating Required Columns](#validating-required-columns)
    * [2.5 Filtering Data](#filtering-data)
    * [2.6 Saving Filtered Data](#saving-filtered-data)
* [4. Usage](#usage)
* [5. Conclusion](#conclusion)

### Script Introduction
+ This Python script is designed to reads data from an Excel file, based on specific criteria. It extracts rows within a specified range and columns containing required information, and  and then copies these filtered records into a new Excel file. 
### Coad Explanation:

#### import Necessary Libraries
```python
     import pandas as pd
```
+ This line imports the Pandas library, which is essential for working with dataframes in Python. Pandas provides functionalities for data manipulation, analysis, and reading/writing various data formats, including Excel files.

#### Input File and Required Columns
```Python
input_file = 'input_file333.xlsx'
required_columns = ["STUDENTNAME", "EMAIL", "PHNO"]
```
##### Here, we define two variables:

##### 
+ input_file: This variable holds the filename of the Excel file containing the student data you want to filter. Make sure to replace 'input_file333.xlsx' with the actual name of your input file.

+ required_columns: This is a list containing the names of the columns you want to extract from the Excel file. In this case, we're interested in student names, emails, and phone numbers.
 
#### Reading the Excel File
``` Python
try:
    file_data = pd.read_excel(input_file)
except FileNotFoundError:
    print("Error: The input file was not found.")
```
+ This code block attempts to read the Excel file using the pd.read_excel function from Pandas. It's wrapped in a try-except block to handle potential errors.

##### 
+ If the file is found, it's read into a Pandas DataFrame named file_data. This DataFrame is a tabular data structure that allows easy manipulation and analysis of your student data.
In case the file is not found, a FileNotFoundError exception is raised, and an error message is printed indicating the issue.

#### Validating Required Columns
```python
missing_columns = []
for col in required_columns:
    if col not in file_data.columns:
        missing_columns.append(col)

if missing_columns:
    raise ValueError(f"Error: Input file '{input_file}' is missing the columns: {', '.join(missing_columns)}")
```
+ This section verifies if the required columns (STUDENTNAME, EMAIL, and PHNO) actually exist in the Excel file you read.

+ It iterates through each column name in required_columns.
If any column name is not found in the DataFrame's columns (file_data.columns), it's added to the missing_columns list.

+ If missing_columns is not empty (meaning some required columns are missing), a ValueError exception is raised with a specific message indicating which columns are missing from the input file.

#### Filtering Data
```python
start_index = 50
end_index = 59

filtered_data = []
for i in range(start_index, end_index + 1):
    filtered_data.append(file_data.loc[i, required_columns])

filtered_data_df = pd.DataFrame(filtered_data)
```

##### Here, we perform the actual data filtering:


+ start_index and end_index define the range of rows (inclusive) you want to extract from the Excel file. You can modify these values based on your specific filtering criteria.
+ An empty list named filtered_data is created to store the filtered rows.
We iterate through the specified row indices using a for loop.
You can modify the script to suit your specific needs. For example:

+ Change the required_columns list to extract different columns.
Adjust the start_index and end_index for different row ranges.
Add additional filtering logic based on other column values.

+ Inside the loop, file_data.loc[i, required_columns] extracts the row
filtered_data_df = pd.DataFrame(filtered_data)
#### Saving Filtered Data
```python
output_file = 'filtered_output.xlsx'

filtered_data_df.to_excel(output_file, index=False)
print("Filtered data has been saved to", output_file)
```
+ output_file: The name of the Excel file to save the filtered data in.
filtered_data_df.to_excel(output_file, index=False): Saves the filtered DataFrame to the output file, excluding the row index column.
A success message is printed upon completion.


#### Error Handling:

The script handles various potential errors, such as:
 * FileNotFoundError: If the input file is not found.
 * ValueError: If the input file is missing any of the required columns.
 * PermissionError: If the script doesn't have permission to access the files.
 * Unexpected errors: General exceptions are caught and printed for debugging.

#### Contributing:
>If you'd like to contribute to this project, please fork the repository, make your changes, and submit a pull request.
### Author :-
- Name : Saisandeep
+ GitHub : https://github.com/saisandeep12345
