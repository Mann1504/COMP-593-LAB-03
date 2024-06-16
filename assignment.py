import pandas as pd
import os
import sys
from datetime import datetime

def validate_input(args):
    if len(args) != 2:
        print("E:\Mann\COMP-593-LAB-03\sales_csv.py.")
        sys.exit(1)
    csv_path = args[1]
    if not os.path.isfile(csv_path):
        print(f"E:\Mann\COMP-593-LAB-03\sales_csv.py: File '{csv_path}' .")
        sys.exit(1)
    return csv_path

def create_orders_directory(base_path):
    today = datetime.today().strftime('%Y-%m-%d')
    orders_dir = os.path.join(base_path, f"Orders_{today}")
    os.makedirs(orders_dir, exist_ok=True)
    return orders_dir

def process_sales_data(csv_path, orders_dir):
    sales_data = pd.read_csv(csv_path)
    order_ids = sales_data['ORDER ID'].unique()
    
    for order_id in order_ids:
        order_data = sales_data[sales_data['ORDER ID'] == order_id]
        order_data = order_data.sort_values(by='ITEM NUMBER')
        order_data['TOTAL PRICE'] = order_data['ITEM QUANTITY'] * order_data['ITEM PRICE']
        
        total_row = pd.DataFrame([['Grand Total', '', '', order_data['TOTAL PRICE'].sum()]], columns=order_data.columns)
        order_data = pd.concat([order_data, total_row])
        
        order_file = os.path.join(orders_dir, f"Order_{order_id}.xlsx")
        with pd.ExcelWriter(order_file, engine='xlsxwriter') as writer:
            order_data.to_excel(writer, index=False, sheet_name='Order')
            worksheet = writer.sheets['Order']
            currency_format = writer.book.add_format({'num_format': '$#,##0.00'})
            worksheet.set_column('D:D', None, currency_format)
            worksheet.set_column('A:D', 18)

def main():
    csv_path = validate_input(sys.argv)
    base_path = os.path.dirname(csv_path)
    orders_dir = create_orders_directory(base_path)
    process_sales_data(csv_path, orders_dir)

if __name__ == "__main__":
    main()
