import pandas as pd
from jinja2 import Environment, FileSystemLoader

# Define the Excel file name and sheet name
excel_file = 'master data.xlsx'  # Change this to your Excel file's name
sheet_name = 'userdata'

# Load the data from the Excel file into a DataFrame
try:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit(1)

# Check if the required columns ("service name" and "user ID") exist in the DataFrame
required_columns = ["Service Name", "User ID"]
if not all(column in df.columns for column in required_columns):
    print("Required columns not found in the Excel sheet.")
    exit(1)

# Load the HTML template
template_loader = FileSystemLoader(searchpath='./')
env = Environment(loader=template_loader)
template = env.get_template('template.html')  # Create an HTML template file (template.html)

# Create a list of dictionaries from the DataFrame
data_list = df.to_dict(orient='records')

# Render the template with the data
rendered_html = template.render(data_list=data_list)

# Save the rendered HTML to a file
with open('output.html', 'w', encoding='utf-8') as output_file:
    output_file.write(rendered_html)

print("HTML file generated successfully.")
