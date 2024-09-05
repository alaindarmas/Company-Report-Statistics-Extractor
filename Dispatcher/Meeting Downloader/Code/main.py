import os
import pandas as pd
import re
import pdfkit
import concurrent.futures
import time

# Function to load configuration values from an Excel file
def load_config(config_path):
    config_df = pd.read_excel(config_path, engine='openpyxl')
    config = {row['Key']: row['Value'] for _, row in config_df.iterrows()}
    config_dir = os.path.dirname(config_path)
    for key in config:
        config[key] = os.path.abspath(os.path.join(config_dir, config[key]))
    return config

# Path to the config file
config_path = r"C:\Users\alain\Alain Developer\Research\Meeting Downloader\Documentation\ConfigFile_MeetingsDownloader.xlsx"
config = load_config(config_path)

# Define required keys for configuration and check for missing keys in the configuration
required_keys = ['path_wkhtmltopdf', 'output_folder_dir']
for key in required_keys:
    if key not in config:
        raise KeyError(f"Missing required configuration key: {key}")

# Paths from the configuration
path_wkhtmltopdf = config['path_wkhtmltopdf']
output_folder_dir = config['output_folder_dir']
 
# Configure pdfkit with the path to wkhtmltopdf
config_pdfkit = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

# Create directories if they don't exist
os.makedirs(output_folder_dir, exist_ok=True)

# Path to the input Excel file
input_excel_path = r"C:\Users\alain\Alain Developer\Research\Meeting Downloader\Input\Initial_Filing_Info.xlsx"

# Verify that the input Excel file exists and is a valid Excel file
if not os.path.exists(input_excel_path):
    raise FileNotFoundError(f"The input Excel file was not found at: {input_excel_path}")

try:
    # Load the input Excel file
    df_input = pd.read_excel(input_excel_path, engine='openpyxl')
except Exception as e:
    raise ValueError(f"Failed to read the input Excel file at {input_excel_path}. Please ensure the file is a valid Excel file. Error: {e}")

# Function to sanitize filenames by replacing invalid characters
def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

# Build a set of existing filenames in the output directory
existing_files = {filename for filename in os.listdir(output_folder_dir) if filename.endswith('.pdf')}

# Function to ensure the URL has a valid scheme
def ensure_valid_url(url):
    if not url.startswith(('http://', 'https://')):
        return 'http://' + url
    return url

# Function to download and save the document as a PDF with the specified naming convention
def download_and_save_document(row):
    url = ensure_valid_url(row["LinkToTxt"])
    ticker = row["Ticker"]
    company_name = row["CompanyName"]
    filed_at = row["FiledAt"]
    fyear = row["FYear"]

    filename = sanitize_filename(f"{ticker}_{company_name}_{filed_at}_{fyear}.pdf")
    file_path = os.path.join(output_folder_dir, filename)

    # Check if the file already exists
    if filename in existing_files:
        print(f"{filename} already exists. Skipping download.")
        return

    try:
        # Convert the web page to a PDF
        pdfkit.from_url(url, file_path, configuration=config_pdfkit)
        print(f'PDF saved successfully at {file_path}')
        # Add the newly downloaded file to the set
        existing_files.add(filename)
    except Exception as e:
        print(f"An error occurred while converting {url} to PDF: {e}")

def format_elapsed_time(elapsed_seconds):
    hours, remainder = divmod(elapsed_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

def main():
    start_time = time.time()
    max_workers = 10  # You can adjust this number based on your CPU capabilities
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(download_and_save_document, row) for index, row in df_input.iterrows()]
        for count, future in enumerate(concurrent.futures.as_completed(futures), 1):
            try:
                future.result()
                if count % 100 == 0:
                    elapsed_time = time.time() - start_time
                    formatted_time = format_elapsed_time(elapsed_time)
                    print(f"Processed {count} files in {formatted_time}")
            except Exception as e:
                print(f"An error occurred: {e}")

# Run the main function if this script is executed
if __name__ == "__main__":
    main()
