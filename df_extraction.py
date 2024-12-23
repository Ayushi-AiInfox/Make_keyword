import pandas as pd


input_file = "Gpt 1 --For Karan--Needs FOCUS.xlsx"
output_file = "demo.xlsx"
print("File Loaded")

data = pd.read_excel(input_file)
demo_data = data.head(100)


demo_data.to_excel(output_file, index=False)

output_file





