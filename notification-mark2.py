import pandas as pd
from jinja2 import Environment, FileSystemLoader
from datetime import datetime

# Define the Excel file name
excel_file = 'master data.xlsx'  # Change this to your Excel file's name

# Load data from the "SL01" and "userdata" sheets into DataFrames
try:
    sl01_df = pd.read_excel(excel_file, sheet_name='SL02', engine='openpyxl')
    userdata_df = pd.read_excel(excel_file, sheet_name='userdata', engine='openpyxl')
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit(1)

# Check if the required columns exist in the DataFrames
if 'Service Name' not in sl01_df.columns:
    print("Required column 'Service Name' not found in the 'SL01' sheet.")
    exit(1)

if 'Service Name' not in userdata_df.columns or 'User ID' not in userdata_df.columns:
    print("Required columns 'Service Name' and 'User ID' not found in the 'userdata' sheet.")
    exit(1)

# Create a dictionary to store the mapping of Service Name to User ID
service_to_user_mapping = {}

# Populate the dictionary by iterating through the "userdata" DataFrame
for index, row in userdata_df.iterrows():
    service_name = row['Service Name']
    user_id = row['User ID']
    service_to_user_mapping[service_name] = user_id

# Load the HTML template from the template.html file
template_loader = FileSystemLoader(searchpath='./')
env = Environment(loader=template_loader)
template = env.get_template('template.html')

# Create a dictionary with Service Name and User ID separated by full stops
data_dict = {}
for service_name in sl01_df['Service Name']:
    user_id = service_to_user_mapping.get(service_name, '')
    data_dict[service_name] = user_id

# Render the template with the data
rendered_html = template.render(data_dict=data_dict)

# Generate a unique file name with a timestamp
build_number = datetime.now().strftime("%Y%m%d%H%M%S")
output_file_name = f"output_{build_number}.html"

# Save the rendered HTML to the unique output file
with open(output_file_name, 'w', encoding='utf-8') as output_file:
    output_file.write(rendered_html)

print(f"HTML file '{output_file_name}' generated successfully.")
