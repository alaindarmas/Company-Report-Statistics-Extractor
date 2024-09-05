import fitz  # PyMuPDF, a library to work with PDF files
import re  # Regular expressions for pattern matching
import os  # Operating system interfaces
import pandas as pd  # Data analysis library
import shutil  # File operations
import nltk  # Natural Language Toolkit for text processing
import time  # Time tracking

# Ensure necessary NLTK data is downloaded
nltk.download('punkt')

# Function to sanitize filenames by replacing invalid characters
def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

# Function to load configuration values from an Excel file
def load_config(config_path):
    config_df = pd.read_excel(config_path, engine='openpyxl')
    config = {row['Key']: row['Value'] for _, row in config_df.iterrows()}
    config_dir = os.path.dirname(config_path)
    for key in config:
        config[key] = os.path.abspath(os.path.join(config_dir, config[key]))
    return config

# Load the configuration from the specified path
config_path = os.path.join(os.path.dirname(__file__), '..', 'Documentation', 'ConfigFile_Performer.xlsx')
config = load_config(config_path)

# Define required keys for configuration and check for missing keys in the configuration
required_keys = ['pdf_folder_dir', 'output_excel_dir', 'input_excel_path']
for key in required_keys:
    if key not in config:
        raise KeyError(f"Missing required configuration key: {key}")

# Paths from the configuration
pdf_folder_dir = config['pdf_folder_dir']
output_excel_dir = config['output_excel_dir']
input_excel_path = config['input_excel_path']

# Create directories if they don't exist
os.makedirs(output_excel_dir, exist_ok=True)

# Verify that the input Excel file exists and is a valid Excel file
if not os.path.exists(input_excel_path):
    raise FileNotFoundError(f"The input Excel file was not found at: {input_excel_path}")

try:
    # Load the input Excel file
    df_input = pd.read_excel(input_excel_path, engine='openpyxl')
except Exception as e:
    raise ValueError(f"Failed to read the input Excel file at {input_excel_path}. Please ensure the file is a valid Excel file. Error: {e}")

# Create a copy of the input Excel file to the output directory
output_excel_path = os.path.join(output_excel_dir, "Meeting_Crawler_Output.xlsx")
shutil.copy(input_excel_path, output_excel_path)

# Load the copied Excel file to add the new columns
df_output = pd.read_excel(output_excel_path, engine='openpyxl')

# Add 'Name Map', 'Nr_of_Meetings_Script', 'Meetings Held Regex', and 'Met Regex' columns
if 'Name Map' not in df_output.columns:
    df_output['Name Map'] = df_output.apply(lambda row: sanitize_filename(f"{row['Ticker']}_{row['CompanyName']}_{row['FiledAt']}_{row['FYear']}.pdf"), axis=1)

if 'Nr_of_Meetings_Script' not in df_output.columns:
    df_output['Nr_of_Meetings_Script'] = pd.Series([None] * len(df_output), dtype='float64')

if 'Meetings Held Regex' not in df_output.columns:
    df_output['Meetings Held Regex'] = pd.Series([None] * len(df_output), dtype='float64')

if 'Met Regex' not in df_output.columns:
    df_output['Met Regex'] = pd.Series([None] * len(df_output), dtype='float64')

# Function to save the DataFrame to the Excel file
def save_to_excel(df, path, start_time):
    with pd.ExcelWriter(path, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False)
    elapsed_time = time.time() - start_time
    print(f"Data saved to {path}. Elapsed time: {elapsed_time:.2f} seconds")

# Function to extract text from a PDF
def extract_text_from_pdf(pdf_path):
    try:
        pdf_document = fitz.open(pdf_path)
        text = ''
        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)
            text += page.get_text()
        return text
    except Exception as e:
        print(f"An error occurred while extracting text from {pdf_path}: {e}")
        return ''
    finally:
        pdf_document.close()

# Function to extract the number of meetings using the first regex
def extract_meetings_held_regex(text):
    match = re.search(r'\b(?:board(?: of directors)?)\s+held\s+(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|\d+)\s+meetings\b', text, re.IGNORECASE)
    if match:
        word_to_number = {
            "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
            "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
            "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, 
            "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18, 
            "nineteen": 19, "twenty": 20
        }
        number_str = match.group(1).lower()
        if number_str.isdigit():
            number_value = int(number_str)
        else:
            number_value = word_to_number.get(number_str, -1)
        return number_value
    else:
        return -1

# Function to extract the number of times the board met using the second regex
def extract_met_regex(text):
    match = re.search(r'\b(?:board(?: of directors)?)\s+met\s+(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|\d+)\s+times\b', text, re.IGNORECASE)
    if match:
        word_to_number = {
            "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
            "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
            "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, 
            "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18, 
            "nineteen": 19, "twenty": 20
        }
        number_str = match.group(1).lower()
        if number_str.isdigit():
            number_value = int(number_str)
        else:
            number_value = word_to_number.get(number_str, -1)
        return number_value
    else:
        return -1

# Function to update the Excel file with the extracted data
def update_excel_with_results(df, index, meetings_held_count, met_count):
    # Determine what to write in the Nr_of_Meetings_Script column
    if meetings_held_count == -1 and met_count == -1:
        final_count = -1
    elif meetings_held_count != -1 and met_count == -1:
        final_count = meetings_held_count
    elif meetings_held_count == -1 and met_count != -1:
        final_count = met_count
    elif meetings_held_count != -1 and met_count != -1:
        if meetings_held_count == met_count:
            final_count = meetings_held_count
        else:
            final_count = -2
    else:
        final_count = -1  # Fallback case

    df.at[index, 'Meetings Held Regex'] = meetings_held_count
    df.at[index, 'Met Regex'] = met_count
    df.at[index, 'Nr_of_Meetings_Script'] = final_count

# Function to process each PDF file
def process_pdf_file(df):
    pdf_files = [f for f in os.listdir(pdf_folder_dir) if f.endswith('.pdf')]
    processed_count = 0
    start_time = time.time()

    for index, row in df.iterrows():
        # Use the 'Name Map' column to match the PDF
        filename = row['Name Map']

        # Check if the file exists in the PDF's folder
        if filename in pdf_files:
            file_path = os.path.join(pdf_folder_dir, filename)
            try:
                # Extract text from the PDF
                text = extract_text_from_pdf(file_path)

                # Apply the regex to extract the number of meetings held
                meetings_held_count = extract_meetings_held_regex(text)

                # Apply the regex to extract the number of times the board met
                met_count = extract_met_regex(text)

                # Update the Excel file with the extracted data
                update_excel_with_results(df, index, meetings_held_count, met_count)

            except Exception as e:
                print(f"An error occurred while processing {filename}: {e}")
        else:
            # If no matching PDF is found, write -1 for all columns
            update_excel_with_results(df, index, -1, -1)
            print(f"No matching PDF found for the row: {filename}. Wrote -1 for this row.")

        processed_count += 1
        print(f"Processed {processed_count} documents.")

        # Save every 50 instances
        if processed_count % 50 == 0:
            save_to_excel(df, output_excel_path, start_time)

    # Save the remaining data after all rows are processed
    save_to_excel(df, output_excel_path, start_time)

def main():
    process_pdf_file(df_output)

# Run the main function if this script is executed
if __name__ == "__main__":
    main()
