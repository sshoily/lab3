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
        print("ERROR: MISSING CSV FILE PATH.")
        sys.exit(1)
    
    # Check whether provided parameter is a valid path of file
    if not os.path.isfile(sys.argv[1]):
        print("ERROR: INVALID CSV FILE PATH.")
        sys.exit(1)
        
    return sys.argv[1]

def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    sales_csv_file = os.path.abspath(sales_csv)
    sales_csv_dir = os.path.dirname(sales_csv_file)
    
    # Determine the name and path of the directory to hold the order data files
    current_date = date.today().isoformat()
    orders_folder = f"orders_{current_date}"
    orders_dir = os.path.join(sales_csv_dir, orders_folder)
    
    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)
    
    return orders_dir

def process_sales_data(sales_csv, orders_dir):
    # Read sales data CSV file
    sales_df = pd.read_csv(sales_csv)

    # Add a new "TOTAL PRICE" column
    sales_df.insert(7, "TOTAL PRICE", sales_df["ITEM QUANTITY"] * sales_df["ITEM PRICE"])

    # Remove columns that are not needed
    sales_df.drop(columns=["ADDRESS", "CITY", "STATE", "POSTAL CODE", "COUNTRY"], inplace=True)

    # Group the rows in the DataFrame by order ID
    for order_id, order_data in sales_df.groupby("ORDER ID"):
        # Remove the "ORDER ID" column
        order_data = order_data.drop(columns=["ORDER ID"])

        # Sort the items by item number
        order_data = order_data.sort_values(by="ITEM NUMBER")
        
        # Append a "GRAND TOTAL" row
        grand_total = order_data["TOTAL PRICE"].sum()
        grand_total_row = pd.DataFrame({
            'ITEM NUMBER': [''],
            'ITEM NAME': [''],
            'ITEM QUANTITY': [''],
            'ITEM PRICE': ['GRAND TOTAL'],
            'TOTAL PRICE': [grand_total]
        })
        order_data = pd.concat([order_data, grand_total_row], ignore_index=True)

        # Determine the file name and full path of the Excel sheet
        file_name = f'ORDER_{order_id}.xlsx'
        file_path = os.path.join(orders_dir, file_name)
        
        # Export the data to an Excel sheet
        order_data.to_excel(file_path, index=False, sheet_name=f'Order {order_id}')
        
        # TODO: Add any additional Excel formatting here, if needed

if __name__ == '__main__':
    main()
