import os
import pandas as pd
import openpyxl as xl


def process_spreadsheet(filepath):
    df = pd.read_excel(filepath, header=None)
    
    # fill missing values with empty string
    df.fillna('', inplace=True)
    
    # get row names that have 'Carrier' in them
    carrier_rows = df[df.iloc[:,0].str.contains('Carrier')].index.to_list()
    carrier_rows.pop(0) # remove the first element
    
    # create a dictionary of carriers
    carriers = {}
    
    for i, start in enumerate(carrier_rows):
        # get the chart row names for the current carrier
        end = carrier_rows[i+1] if i+1 < len(carrier_rows) else len(df)
        carriers[df.iloc[start,0].strip("Carrier:  ")] = df.iloc[start:end]
    
    # create a dictionary of charts per carrier from carriers
    charts = {}
    for carrier in carriers.keys():
        charts[carrier] = {}
        df = carriers[carrier]
        chart_rows = df[df.iloc[:,1].str.contains('Chart #')].index.to_list()
        for i, start in enumerate(chart_rows):
            end = chart_rows[i+1]-1 if i+1 < len(chart_rows) else df.index[-1]
            charts[carrier][df.loc[start,1].strip()] = df.loc[start:end].reset_index(drop=True)
    
    carriers_line = {}
    for carrier in carriers.keys():
        carriers_line[carrier] = carriers[carrier].iloc[:2]
        
    return charts, carriers_line



def write_carrier(carriers, new_filepath, sheet, carrier):
    with pd.ExcelWriter(new_filepath, engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
            # Read only the specific sheet you want to append to
        book = writer.book
        try:
            startrow = book[sheet].max_row
        except KeyError:
            startrow = 0
        carriers[carrier].to_excel(writer, sheet_name=sheet, startrow=startrow, header=False, index=False)
        
    wb = xl.load_workbook(new_filepath)
    ws = wb[sheet]

        # Define the font
    font1 = xl.styles.Font(name='Tahoma', size=8, bold=True)
    font2 = xl.styles.Font(name='Tahoma', size=7, bold=True)

    ws.merge_cells(start_row=startrow+1, start_column=1, end_row=startrow+1, end_column=3)
    ws.merge_cells(start_row=startrow+1, start_column=4, end_row=startrow+1, end_column=6)
    for row in ws.iter_rows(min_row=startrow+1, max_row=startrow+3, min_col=1, max_col=12):
        for cell in row:
            cell.font = font1 if cell.column < 8 else font2
            if cell.row == startrow+1:
                if cell.column < 8:
                    cell.number_format = 'General'
                else:
                    cell.number_format = '$#,##0.00'
            else:
                cell.number_format = '0.00%'
                
    wb.save(new_filepath)


def write_charts(charts, new_filepath, sheet, carrier):
    with pd.ExcelWriter(new_filepath, engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
            # Read only the specific sheet you want to append to
        book = writer.book
        try:
            startrow = book[sheet].max_row
        except KeyError:
            startrow = 0

        df = pd.concat(charts[carrier].values(), ignore_index=True)
        chart_rows = df[df.iloc[:,1].str.contains('Chart #')].index

        df.to_excel(writer, sheet_name=sheet, startrow=startrow, header=False, index=False)
        
    wb = xl.load_workbook(new_filepath)
    ws = wb[sheet]

        # Define the font
    font1 = xl.styles.Font(name='Tahoma', size=7, bold=True)
    font2 = xl.styles.Font(name='Tahoma', size=7, bold=False)
    for row in ws.iter_rows(min_row=startrow+1, max_row=startrow+df.shape[0], min_col=1, max_col=12):
        for cell in row:
            if cell.row in chart_rows+startrow+1:
                cell.font = font1
            else:
                cell.font = font2
                cell.number_format = '$#,##0.00'
                
    wb.save(new_filepath)



new_charts, new_carriers = process_spreadsheet('Spreadsheet_New.xlsx')
old_charts, old_carriers = process_spreadsheet('Spreadsheet_Old.xlsx')

for carrier in new_charts.keys():
    if carrier not in old_charts.keys():
        print(f'New carrier "{carrier}"')
        old_carriers[carrier] = new_carriers[carrier]
        old_charts[carrier] = new_charts[carrier]
    else:
        for chart in new_charts[carrier].keys():
            if chart not in old_charts[carrier].keys():
                print(f'New chart for carrier "{carrier}": {chart}')
                old_charts[carrier][chart] = new_charts[carrier][chart]



new_filepath = 'Spreadsheet_Updated.xlsx'
sheet = 'CarrierArDetail'
# Check if the file exists
if not os.path.exists(new_filepath):
    # If the file doesn't exist, create an empty DataFrame and write it to the file
    pd.DataFrame().to_excel(new_filepath, sheet_name=sheet)
for carrier in old_charts.keys():
    print(f'Writing "{carrier}"...')
    write_carrier(old_carriers, new_filepath, sheet, carrier)
    write_charts(old_charts, new_filepath, sheet, carrier)