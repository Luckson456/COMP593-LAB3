""" 
Description: 
Divides sales data CSV file into individual order data Excel files.

Usage:
python process_sales_data.py sales_csv_path

Parameters:
sales_csv_path = Full path of the sales data CSV file
"""
import os 
import sys
from datetime import date
import re 
import pandas as pd 

def main():
    sales_csv_path = get_sales_csv_path()
    orders_dir_path = create_orders_dir(sales_csv_path)
    process_sales_data(sales_csv_path, orders_dir_path)

def get_sales_csv_path():
    """Gets the path of sales data CSV file from the command line

    Returns:
        str: Path of sales data CSV file
    """
    # TODO: Check whether command line parameter provided
    num_params =len(sys.argv)-1
    if num_params< 1:
        print('Error: Missing path to sales data CSV file')
        sys.exit(1)
        
    # TODO: Check whether provide parameter is valid path of file
    sales_csv_path =sys.argv[1]
    if not os.path.isfile(sales_csv_path):
        print('Error :Invalid path to sales data CSV file')
        sys.exit(1)
        
    # TODO: Return path of sales data CSV file
        return  sales_csv_path

def create_orders_dir(sales_csv_path):
    """Creates the directory to hold the individual order Excel sheets

    Args:
        sales_csv_path (str): Path of sales data CSV file

    Returns:
        str: Path of orders directory
    """
    # TODO: Get directory in which sales data CSV file resides
    sales_dir_path= os.path.dirname(os.path.abspath(sales_csv_path))
    # TODO: Determine the path of the directory to hold the order data files
    today_date = date.today().isoformat()
    orders_dir_path =os.path.join(sales_csv_path)
    # TODO: Create the orders directory if it does not already exist
    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)
    # TODO: Return path of orders directory
    return orders_dir_path

def process_sales_data(sales_csv_path, orders_dir_path):
    """Splits the sales data into individual orders and save to Excel sheets

    Args:
        sales_csv_path (str): Path of sales data CSV file
        orders_dir_path (str): Path of orders directory
    """
    # TODO: Import the sales data from the CSV file into a DataFrame
    sales_df =pd.read_csv(sales_csv_path)
    # TODO: Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, 'Total Price',sales_df['Item Quantity']*sales_df['Item Price'])
    # TODO: Remove columns from the DataFrame that are not needed
    sales_df.drop(columns=['ADDRESS','CITY','STATE','POSTAL CODE','COUNTRY'],inplace =True)
    # TODO: Groups orders by ID and iterate 
    for order_id,order_df in sales_df.groupby('ORDER ID'):
        # TODO: Remove the 'ORDER ID' column
        order_df =order_df.drop(columns=['ORDER ID'])
        # TODO: Sort the items by item number
        order_df=order_df.sort_values('Item Number')
        # TODO: Append a "GRAND TOTAL" row
        grand_total =order_df['Total Price'].sum()
        grand_total_row=pd.DataFrame({'Item number':['Grand Total'],'Item Quantity':['-'],'Item Price':['-'],'Total Price':[grand_total]})
        order_df=pd.concat([order_df,grand_total_row],ignore_index=True)
        # TODO: Determine the file name and full path of the Excel sheet
        order_file_name =f"order_{order_id}.xlsx"
        order_file_path =os.path.join(orders_dir_path,order_file_name)     
        # TODO: Export the data to an Excel sheet
        with pd.ExcelWriter(order_file_path) as writer:
            order_df.to_excel(writer,index=False, sheet_name=f"Order{order_id}")
        # TODO: Format the Excel sheet
        worksheet=writer.sheets[f"Order{order_id}"]
        # TODO: Define format for the money columns
        money_format ='${:,2f}'
        # TODO: Format each column
        worksheet.set_column('C:D',15,writer.book.add_format({'num_format':money_format}))
        worksheet.set_column('E:E',15,writer.book .add_format({'num_format':money_format}))
        # TODO: Close the sheet
        writer.close()

    return

if __name__ == '__main__':
    main()