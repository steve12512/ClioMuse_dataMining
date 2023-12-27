import pandas as pd

# Load the Excel file
file_path = 'Booking Stats.xlsx'
xlsx = pd.ExcelFile(file_path)

# Load the 'codes & prices' sheet into a DataFrame
codes_prices_df = pd.read_excel(file_path, sheet_name='Codes & Prices')

# Define a function to get the price from 'codes & prices' sheet
def get_price_from_codes(code, month, codes_prices_df):
    # Find the column index for the month
    month_col_index = codes_prices_df.columns.get_loc(month)
    
    # Find the row index for the product code
    code_row_index = codes_prices_df.index[codes_prices_df['Product Code'] == code].tolist()
    
    # If the product code is found, return the price for that month
    if code_row_index:
        return codes_prices_df.iloc[code_row_index[0], month_col_index]
    else:
        return None

# Iterate over each sheet name
for sheet_name in xlsx.sheet_names:
    if sheet_name in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']:
        # Load the sheet into a DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Convert columns to numeric, coercing errors to NaN
        df['net_price'] = pd.to_numeric(df['net_price'], errors='coerce')
        df['retail_price'] = pd.to_numeric(df['retail_price'], errors='coerce')
        df['num_of_travellers'] = pd.to_numeric(df['num_of_travellers'], errors='coerce')

        # Iterate through each row to check for missing net and retail prices
        for index, row in df.iterrows():
            net_price = row['net_price']
            retail_price = row['retail_price']
            num_travelers = row['num_of_travellers']
            code_of_product = row['product_code']

            # Check and calculate missing prices
            if pd.isna(net_price) and not pd.isna(retail_price) and num_travelers > 0:
                # Calculate net price from retail price
                df.at[index, 'net_price'] = retail_price / num_travelers
            elif not pd.isna(net_price) and pd.isna(retail_price):
                # Calculate retail price from net price
                df.at[index, 'retail_price'] = net_price * num_travelers
            elif pd.isna(net_price) and pd.isna(retail_price):
                # Look up the price using the code of product and the name of the month as the column
                found_price = get_price_from_codes(code_of_product, sheet_name, codes_prices_df)
                if found_price is not None:
                    df.at[index, 'net_price'] = found_price
                    df.at[index, 'retail_price'] = found_price * num_travelers


        # Save the updated DataFrame back to the Excel file
        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Price filling complete.")
