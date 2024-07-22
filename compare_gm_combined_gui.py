import pandas as pd
import openpyxl
from datetime import datetime
import tkinter as tk
from tkinter import filedialog as fd 
from tkinter import simpledialog as sd
import os 

# Functions
def main_gm():
  current_date = datetime.now().strftime('%Y-%m-%d') # Current date for naming purposes

  # User input for file path  
  radley_status_inquiry, forecast_13_week = get_paths()

  # Pass file paths to be read into pandas dfs
  df1,df2 = read_excel(radley_status_inquiry, forecast_13_week)

  # Cleans and aligns dfs
  df1 = process_status_inquiry(df1)
  df2 = process_13week(df2)
  
  # Compares, styles, and exports dfs
  compare_files(df1, df2)

def get_paths(): 
  radley_status_inquiry = fd.askopenfile(title = 'Open Status Inquiry for One Company')
  forecast_13_week = fd.askopenfile(title = 'Open 13-Week')

  if radley_status_inquiry and forecast_13_week:
    radley_path = os.path.abspath(radley_status_inquiry.name)
    forecast_13_week_path = os.path.abspath(forecast_13_week.name)
  else: 
    print("Please Select Two Files")
  
  return radley_path, forecast_13_week_path

def read_excel(radley_status_inquiry, forecast_13_week):
  df1 = pd.read_excel(radley_status_inquiry, engine = "openpyxl")
  df2 = pd.read_excel(forecast_13_week, skiprows=[0,1,2,3], engine = "openpyxl")

  df1['Part Number'] = df1['Part Number'].astype(str)
  df2['Part #'] = df2['Part #'].astype(str)

  return df1, df2

# RADLEY STATUS INQUIRY
def process_status_inquiry(df1):
  # Converts date to datetime
  df1.loc[:,'Date'] = pd.to_datetime(df1['Date'].dt.date)

  # Renames the columns that will be compared later
  df1.rename(columns={'Div': 'Company', 'Part Number': 'Part #'}, inplace=True)

  # Sets a multi-index so that when the table is pivoted the company is garuanteed to stay associated with the part number
  df1.set_index(['Company','Part #'], inplace = True)
  df1 = df1.pivot_table(columns='Date', values='Net', index=['Company', 'Part #'], aggfunc='sum', fill_value= 'Part # Missing Date')

  # Flatten data frame to move company and part number out of index and forward fill company so it associates appropriately
  df1.columns = pd.to_datetime(df1.columns)
  df1 = df1.reset_index()
  df1['Company'] = df1['Company'].ffill()

  date_headers = []
  for col in df1.columns:
      try:
          # Attempt to convert the header to a datetime object
          pd.to_datetime(col)
          # If successful, add to the list
          date_headers.append(col)
      except ValueError:
          # If conversion fails, continue
          pass

  for col in date_headers:
    df1.fillna(0, inplace=True)

  return df1

# 13 WEEK FORECAST
# Create mask to remove all but companies and part numbers
def process_13week(df2):
  import re

  part_number_pattern = r'^[A-Z0-9]+$' # Part num regex pattern
  company_pattern = r'^[A-Za-z]+$' # Company regex pattern

  # Drop blank rows
  df2 = df2.dropna(how='all')
  df2.reset_index(drop = True, inplace=True) # Resets index

  # MASKS
  company_mask = df2.iloc[:,0].str.match(company_pattern) # Create a mask that matches the regex pattern
  company_mask = company_mask.fillna(False) # Fill in missing values with False
  non_company_mask = ~company_mask

  part_number_mask = df2.iloc[:,0].str.match(part_number_pattern) # Create a mask that matches the regex pattern
  part_number_mask = part_number_mask.fillna(False) # Fill in missing values with False
  non_part_number_mask = ~part_number_mask

  # Apply the mask to get only the part numbers
  part_numbers = df2[part_number_mask]

  # Apply the mask to get only the part numbers
  companies = df2[company_mask]

  # Apply Masks
  df2 = df2[part_number_mask] # Isolates part no and company
  df2['Company'] = df2.iloc[:,0].mask(non_company_mask)
  df2['Company'] = df2['Company'].ffill()
  df2 = df2.loc[non_company_mask] # Removes companies from Part Numbers

  # Drop columns unecessary for comparison
  df2.drop(columns=['Series', 'Kanban', 'Description', 'Color'], inplace=True)

  # Flatten data frame to move company and part number out of index and forward fill company so it associates appropriately
  df2 = df2.reset_index()
  df2['Company'] = df2['Company'].ffill()

  return df2

def focus_company(df2):
  # Input for which company is being checked
  focus_company = sd.askstring('Company', 'Which company would you like to compare?')

  # Cleans user input 
  focus_company = focus_company.upper()
  focus_company = focus_company.strip()
  df2 = df2[df2["Company"] == focus_company]

  return focus_company

def compare_files(df1, df2):
  f_company = focus_company(df2)

  current_date = datetime.now().strftime('%Y-%m-%d')
  df1.set_index(['Company', 'Part #'], inplace = True)
  df2.set_index(['Company', 'Part #'], inplace = True)

  # Aligning DataFrames on common columns
  common_columns = df2.columns.intersection(df1.columns)
  common_indices = df2.index.intersection(df1.index)

  df1_aligned = df1.loc[common_indices, common_columns]
  df2_aligned = df2.loc[common_indices, common_columns]

  # Initialize DataFrame for differences
  df_differences = pd.DataFrame(index=common_indices, columns=common_columns)

  # Find differences
  for index in common_indices:
    for column in common_columns:
      if df1_aligned.loc[index, column] != 'Part # Missing Date' and df2_aligned.loc[index, column] != 'Part # Missing Date':
        if df2_aligned.loc[index, column] != df1_aligned.loc[index, column]:
          df_differences.loc[index, column] =  ('Diff ' + 'RadStatIn: '+ str(df1_aligned.loc[index, column])+ ' 13W: '+str(df2_aligned.loc[index, column]))
        else:
          df_differences.loc[index, column] = ('Match ' + 'RadStatIn: '+ str(df1_aligned.loc[index, column])+ ' 13W: '+str(df2_aligned.loc[index, column]))

  # Rows in df1 but not in df2
  df_extra_in_df1 = df1[~df1.index.isin(df2.index)]

  # Fill NaNs with False for clarity (optional)
  df_differences = df_differences.fillna('Date not shared')

  styled_df = df_differences.style.map(highlight_match).map(highlight_diff).map(highlight_no_comp)

  styled_df.to_excel(f'{current_date} {f_company} Differences.xlsx')

  return

def highlight_match(x):
    green = '#8BC58C'
    match_style = 'background-color: {}; color: white;'.format(green)
    return match_style if 'Match' in x else ''

def highlight_diff(x):
    red = '#C34A4A'
    no_match_style = 'background-color: {}; color: white;'.format(red)
    return no_match_style if 'Diff' in x else ''

def highlight_no_comp(x):
    grey = '#A9A9A9'
    no_comp_style = 'background-color: {}; color: white;'.format(grey)
    return no_comp_style if 'No Comparison' in x else ''

if __name__ == '__main__':
  main()