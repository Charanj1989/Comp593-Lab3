import sys
import os
from datetime import date
import pandas as pd

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

def get_sales_csv():
    if len(sys.argv) < 2:
        print("Error: Missing the parameter for CSV filepath")
        sys.exit(1)
    sales_data_path = sys.argv[1]
    if not os.path.isfile(sales_data_path):
        print("Error: Provided file is not existing on the system....")
        sys.exit(2)
    return sales_data_path

def create_orders_dir(sales_csv):
    sales_dir_path = os.path.dirname(os.path.abspath(sales_csv))
    todays_date = date.today().isoformat()
    order_dir_name = f'orders_{todays_date}'
    orders_dir_path = os.path.join(sales_dir_path, order_dir_name)
    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)
    return orders_dir_path

def process_sales_data(sales_csv, orders_dir):
    sales_data = pd.read_csv(sales_csv)

    sales_data['TOTAL PRICE'] = sales_data['ITEM QUANTITY'] * sales_data['ITEM PRICE']

    sales_data.drop(['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], axis=1, inplace=True)

    group_sales_data = sales_data.groupby('ORDER ID')

    for id, sales_data2 in group_sales_data:
        sales_data2 = sales_data2.copy()
        sales_data2.drop(['ORDER ID'], axis=1, inplace=True)

        sales_data2.sort_values(by='ITEM NUMBER', inplace=True)

        total = sales_data2['TOTAL PRICE'].sum()
        grand_total_row = pd.DataFrame([{'ITEM PRICE': 'GRAND TOTAL:', 'TOTAL PRICE': f"${total:.2f}"}])
        sales_data2 = pd.concat([sales_data2, grand_total_row], ignore_index=True)

        file_name = f"{id}.xlsx"
        proper_file_path = os.path.join(orders_dir, file_name)

        with pd.ExcelWriter(proper_file_path, engine="xlsxwriter") as writer:
            sales_data2.to_excel(writer, index=False, sheet_name='Dataofsales')
            workbook = writer.book
            worksheet = writer.sheets['Dataofsales']
            format_currency = workbook.add_format({"num_format": "$#,##0.00"})
            worksheet.set_column('A:A', 11)
            worksheet.set_column('B:B', 13)
            worksheet.set_column('C:C', 15)
            worksheet.set_column('D:D', 15)
            worksheet.set_column('E:E', 15)
            worksheet.set_column('F:F', 13, format_currency)
            worksheet.set_column('G:G', 13, format_currency)

if __name__ == '__main__':
    main()
