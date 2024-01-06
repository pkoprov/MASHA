import os
import pandas as pd
import openpyxl as xl
import tkinter as tk
from tkinter import filedialog



# def write_carrier(carriers, new_filepath, sheet, carrier):
#     with pd.ExcelWriter(new_filepath, engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
#         # Read only the specific sheet you want to append to
#         book = writer.book
#         try:
#             startrow = book[sheet].max_row
#         except KeyError:
#             startrow = 0
#         carriers[carrier].to_excel(writer, sheet_name=sheet, startrow=startrow, header=False, index=False)
        
#     wb = xl.load_workbook(new_filepath)
#     ws = wb[sheet]

#     # Define the font
#     font1 = xl.styles.Font(name='Tahoma', size=8, bold=True)
#     font2 = xl.styles.Font(name='Tahoma', size=7, bold=True)

#     ws.merge_cells(start_row=startrow+1, start_column=1, end_row=startrow+1, end_column=3)
#     ws.merge_cells(start_row=startrow+1, start_column=4, end_row=startrow+1, end_column=6)
#     col = None
#     typ = None 
#     if 'new' in carriers.keys() and carrier in carriers['new']:
#         col = 'FFFF00'
#         typ = 'solid'
#     fill_col = xl.styles.PatternFill(start_color=col, end_color=col, fill_type=typ)
        
#     for row in ws.iter_rows(min_row=startrow+1, max_row=startrow+3, min_col=1, max_col=12):
#         for cell in row:
#             cell.font = font1 if cell.column < 8 else font2
#             if cell.row == startrow+1:
#                 if cell.column < 8:
#                     cell.number_format = 'General'
#                 else:
#                     cell.number_format = '$#,##0.00'
#             else:
#                 cell.number_format = '0.00%'
#             cell.fill = fill_col
                
#     wb.save(new_filepath)


# def write_charts(charts, new_filepath, sheet, carrier):
#     with pd.ExcelWriter(new_filepath, engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
#             # Read only the specific sheet you want to append to
#         book = writer.book
#         try:
#             startrow = book[sheet].max_row
#         except KeyError:
#             startrow = 0
#         new = charts[carrier].pop('new', [])
#         removed = charts[carrier].pop('removed', [])
#         df = pd.concat(charts[carrier].values(), ignore_index=True)
#         chart_rows = df[df.iloc[:,1].str.contains('Chart #')].index

#         df.to_excel(writer, sheet_name=sheet, startrow=startrow, header=False, index=False)
        
#     wb = xl.load_workbook(new_filepath)
#     ws = wb[sheet]

#     # Define the font
#     font1 = xl.styles.Font(name='Tahoma', size=7, bold=True)
#     font2 = xl.styles.Font(name='Tahoma', size=7, bold=False)
#     col = None
#     typ = None

#     for row in ws.iter_rows(min_row=startrow+1, max_row=startrow+df.shape[0], min_col=1, max_col=12):
#         # check if the chart is new and choose infill color
#         chart = row[1].value.strip() if row[1].value is not None else ''

#         if "Chart #" in chart:
#             if chart in new:
#                 col = 'FFFF00'
#                 typ = 'solid'
#             elif chart in removed:
#                 col = 'FF0000'
#                 typ = 'solid'
#             else:
#                 col = None
#                 typ = None
#         fill_col = xl.styles.PatternFill(start_color=col, end_color=col, fill_type=typ)
#         for cell in row:
#             if cell.row in chart_rows+startrow+1:
#                 cell.font = font1  
#             else:
#                 cell.font = font2
#                 cell.number_format = '$#,##0.00'
#             cell.fill = fill_col
                
#     wb.save(new_filepath)



def select_file(init_dir=None, title=None):
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename(initialdir=init_dir, title=title)
    root.destroy()
    return filepath

def key_plus_df(df, rows, iter, start, col=0):
    end = rows[iter+1] if iter+1 < len(rows) else df.index[-1]
    key = df.loc[start,col].strip()
    df = df.loc[start:end-1]
    return key, df

def process_spreadsheet(filepath):
    df = pd.read_excel(filepath, header=None)
    
    # fill missing values with empty string
    df.fillna('', inplace=True)
    
    # get row names that have 'Carrier' in them
    carrier_rows = df[df.iloc[:,0].str.contains('Carrier')].index.to_list()
    carrier_rows.pop(0) # remove the first element
    
    # create a dictionary of carriers
    charts = {}
    
    for i, car_start in enumerate(carrier_rows):
        # get the chart row names for the current carrier
        carrier, car_df = key_plus_df(df, carrier_rows, i, car_start)
        print("Working with", carrier, "...")

        chart_rows = car_df[car_df.iloc[:,1].str.contains('Chart #')].index.to_list()
        charts[carrier] = {"Carrier info": df.iloc[car_start:chart_rows[0]]}
        for j, ch_start in enumerate(chart_rows):
            chart, ch_df = key_plus_df(car_df, chart_rows, j, ch_start,1)
            print(5*"\t", chart,"...")
            
            visit_rows = ch_df[ch_df.iloc[:,0].str.contains('Aging Date:')].index.to_list()
            charts[carrier][chart] = {"Chart info": df.iloc[ch_start:visit_rows[0]]}
            for k, v_start in enumerate(visit_rows):
                date, v_df = key_plus_df(ch_df, visit_rows, k, v_start)
                print(7*"\t", date)

                charts[carrier][chart][date] = v_df.reset_index(drop=True)
        
    return charts



print('Select the new spreadsheet')
new = select_file(title = 'Select the new spreadsheet')
print('Select the old spreadsheet')
old = select_file(title='Select the old spreadsheet')

new_charts = process_spreadsheet(new)
old_charts = process_spreadsheet(old)

df = pd.DataFrame(columns = ["state"])
for carrier in new_charts.keys():
    if carrier not in old_charts.keys():
        df = pd.concat([df,new_charts[carrier]['Carrier info']], ignore_index=True)
        car_dict = [chart for key,chart in new_charts[carrier].items() if 'Carrier info' not in key][0]
        char_df = pd.concat(list(car_dict.values()),ignore_index=True)
        df = pd.concat([df,char_df],ignore_index=True)
        df["state"].fillna('new', inplace=True)
        df.loc[len(df)] = ''
 
        
    else:
        charts  = set(new_charts[carrier].keys()) | set(old_charts[carrier].keys())
        car_df = new_charts[carrier]['Carrier info']
        df = pd.concat([df,car_df])
        df["state"].fillna('old', inplace=True)

        for chart in charts:
            if chart == 'Carrier info':
                continue
            if chart not in old_charts[carrier].keys():
                print(f'New {chart} for "{carrier}"')
                chart_df = pd.concat(list(new_charts[carrier][chart].values()),ignore_index=True)
                df = pd.concat([df,chart_df])
                df["state"].fillna('new', inplace=True)
            elif chart not in new_charts[carrier].keys():
                print(f'{chart} for "{carrier}" was removed')
                chart_df = pd.concat(list(old_charts[carrier][chart].values()),ignore_index=True)
                df = pd.concat([df,chart_df])
                df["state"].fillna('removed', inplace=True)
            else:
                print(f'{chart} for "{carrier}" was not changed')
                chart_df = pd.concat(list(new_charts[carrier][chart].values()),ignore_index=True)
                df = pd.concat([df,chart_df])
                df["state"].fillna('old', inplace=True)
            df.loc[len(df)] = ''
                


new_filepath = 'Spreadsheet_Updated.xlsx'
sheet = 'CarrierArDetail'
# Check if the file exists
if not os.path.exists(new_filepath):
    # If the file doesn't exist, create an empty DataFrame and write it to the file
    pd.DataFrame().to_excel(new_filepath, sheet_name=sheet)

df.to_excel(new_filepath, sheet_name=sheet, startrow=0, header=False, index=False)

wb = xl.load_workbook(new_filepath)
ws = wb[sheet]


for row in ws.iter_rows(min_row=1, max_row=df.shape[0]):
    if row[0].value == 'new':
        col = 'FFFF00'
        typ = 'solid'
    elif row[0].value == 'removed':
        col = 'FF0000'
        typ = 'solid'
    else:
        col = None
        typ = None
    for cell in row:
        fill_col = xl.styles.PatternFill(start_color=col, end_color=col, fill_type=typ)
        cell.fill = fill_col

wb.save(new_filepath)