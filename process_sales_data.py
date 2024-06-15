""" 
Description: 
  Divides sales data CSV file into individual order data Excel files.

Usage:
  python process_sales_data.py sales_csv_path

Parameters:
  sales_csv_path = Full path of the sales data CSV file
"""

import sys
import os
from datetime import date
import pandas as pd

def main():
    path_sales_csv = get_path_sales_csv()
    orders_dir = create_orders_dir(path_sales_csv)
    process_sales_data(path_sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_path_sales_csv():
    # Check whether command line parameter provided
    if len(sys.argv) < 2:
        print("Error: Missing the parameter for CSV filepath")
        sys.exit(2)

     # Check whether provide parameter is valid path of file
    path_csv = sys.argv[1]
    if not os.path.isfile(path_csv):
        print("ERROR: Invalid CSV file path.")
        sys.exit(2)
    return path_csv

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(path_sales_csv):
    """Creates the directory to hold the individual order Excel sheets

    Args:
        path_sales_csv (str): Path of sales data CSV file

    Returns:
        str: Path of orders directory
    """

    
# Get directory in which sales data CSV file resides
    sales_dir_path = os.path.dirname(os.path.abspath(path_sales_csv))
    
    # Determine the name and path of the directory to hold the order data files
    todays_date = date.today().isoformat()
    order_dir_name = f'orders_{todays_date}'
    orders_dir_path = os.path.join(sales_dir_path,order_dir_name)
    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)
    return orders_dir_path

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(path_sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_data = pd.read_csv('sales_data.csv')
    # Insert a new "TOTAL PRICE" column into the DataFrame
    new_list = sales_data['ITEM PRICE']
    new_list_1 = list(new_list)
    def multiply(a):
        for g in new_list_1:
            c = g*a
            new_list_1.remove(g)
            return c
            break
    sales_data.insert(7,'TOTAL PRICE',[multiply(a) for a in sales_data['ITEM QUANTITY']] )
    # Remove columns from the DataFrame that are not needed
    sales_data.drop(['ADDRESS','CITY','STATE','POSTAL CODE','COUNTRY'],axis=1,inplace=True)
    # Group the rows in the DataFrame by order ID
    group_sales_data = sales_data.groupby('ORDER ID')
    
    
    # For each order ID:
    for id,sales_data2 in group_sales_data:
        
        # Remove the "ORDER ID" column
        sales_data2.drop(['ORDER ID'],axis=1,inplace=True)
        
        #Sort the items by item number
        sales_data2.sort_values(by = 'ITEM NUMBER',inplace = True)
    
        # Append a "GRAND TOTAL" row
        a = sum(sales_data2['TOTAL PRICE'])
        
        new_value = f"${a}"
        new_row = {'ITEM PRICE':'GRAND TOTAL:','TOTAL PRICE': new_value}
        sales_data2.loc[len(sales_data2)] = new_row
        
        # Determine the file name and full path of the Excel sheet
        file_path = os.path.abspath(orders_dir)
        file_name = f"{id}.xlsx"
        proper_file_path = os.path.join(file_path,file_name)
        # Export the data to an Excel sheet
        sales_data2.to_excel(proper_file_path,index=False,sheet_name='Dataofsales')
        # TODO: Format the Excel sheet
        writer = pd.ExcelWriter(proper_file_path, engine="xlsxwriter")
        sales_data2.to_excel(writer,index=False, sheet_name='Dataofsales')
        workbook = writer.book
        worksheet = writer.sheets['Dataofsales']
        format = workbook.add_format({"num_format": "$#,##0.000"})
        worksheet.set_column('A:A',11)
        worksheet.set_column('B:B',13)
        worksheet.set_column('C:C',15)
        worksheet.set_column('D:D',15)
        worksheet.set_column('E:E',15)
        worksheet.set_column('F:F',13,format)
        worksheet.set_column('G:G',13,format)
        worksheet.set_column('H:H',10)
        worksheet.set_column('I:I',30)
        writer.close()
        
    pass

if __name__ == '__main__':
    main()