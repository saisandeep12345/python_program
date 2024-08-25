import pandas as pd

# Initialize the input file name
input_file = 'input_files.xlsx'

try:
    # Read the Excel file using pandas 'read_excel' function
    file_data = pd.read_excel(input_file)

    # Required columns
    required_columns = ["STUDENTNAME", "EMAIL", "PHNO"]
    
    # Validate if the required columns exist in the file
    missing_columns = []
    for col in required_columns:
        if col not in file_data.columns:
            missing_columns.append(col)
    
    if missing_columns :
        raise ValueError(f"Error: Input file '{input_file}' is missing the columns: {', '.join(missing_columns)}")

    # Define the start and end index for rows to be filtered
    start_index = 50
    end_index = 59

    # Filter the data between the specified row indices
    filtered_data = []
    for i in range(start_index, end_index + 1):
        filtered_data.append(file_data.loc[i, required_columns])

    # Convert the list of rows to a DataFrame
    filtered_data_df = pd.DataFrame(filtered_data)

    # Define the output file name
    output_file = 'filtered_output1.xlsx'

    # Save the filtered data to a new Excel file
    filtered_data_df.to_excel(output_file, index=False)
    print("Filtered data has been saved to this New File : {}".format(output_file))

except FileNotFoundError:
  print("This Errors raises the Input file not found here.....")
except ValueError:
  print("This Error can raise the Input file does not contain all required columns as per your requirements.")
except PermissionError:
    print("the Error of Permission denied.so, Please check file permissions.")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
