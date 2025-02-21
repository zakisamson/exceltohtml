import pandas as pd
from jinja2 import Environment, FileSystemLoader

# Load the Excel file
df = pd.read_excel("./data/data.xlsx", sheet_name="SalesOrders")

# Preview the data
print(df.head())

# Load the Jinja2 template
env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("./templates/template.html")

# Convert DataFrame to a list of lists
html_output = template.render(columns=df.columns, data=df.values.tolist())

# Save as an HTML file
with open("./output/output.html", "w", encoding="utf-8") as file:
    file.write(html_output)

print("HTML file generated successfully! 🎉")
