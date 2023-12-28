#import what you have to
import pandas as pd
import matplotlib as plt
import os

#SET THE METHODS WE WILL BE USING

def read_files():
    #read dataframes from the excel files

    try:
        if os.path.exists('dataframe1.xlsx'):
            dataframe1 = pd.read_excel('dataframe1.xlsx')
        else:
            dataframe1 = combine_review_sheets()
            
        if os.path.exists('dataframe2.xlsx'):
            dataframe2 = pd.read_excel('dataframe2.xlsx')
        else:
            dataframe2 = combine_booking_sheets()
            
    except Exception as e:
        print(f"An error occurred: {e}")
    return dataframe1, dataframe2

def combine_review_sheets():
    #read and instantiate dataframe

    # File path to your Excel file
    file_path = 'reviews data.xlsx'  # Replace with your actual file path

    #load Excel file
    xlsx = pd.ExcelFile(file_path)

    #collect all unique column titles from each sheet
    all_columns = []

    #iterate through each sheet to collect column names
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1, nrows=0)  # Read only the second row for column names
        all_columns.extend([col for col in df.columns if col not in all_columns and not 'Unnamed' in str(col)])

    #initialize an empty DataFrame to store combined data
    combined_df = pd.DataFrame()

    #iterate through each sheet
    for sheet_name in xlsx.sheet_names:
        # Read the sheet into a DataFrame, starting from the second row for data
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)

        # Iterate through each row to find the first entirely empty row
        for i in range(len(df)):
            if pd.isna(df.iloc[i]).all():  # Check if all elements in the row are NaN
                break
        # Keep only the data above the first entirely empty row
        df = df.iloc[:i]

        # Add a column to indicate the source sheet
        df['Source Sheet'] = sheet_name
        
        # Append this DataFrame to the combined DataFrame
        combined_df = pd.concat([combined_df, df], ignore_index=True, sort=False)

    # Reindex the combined DataFrame to include all collected columns plus the Source Sheet column
    combined_df = combined_df.reindex(columns=all_columns + ['Source Sheet'])

    #in the first column of combined_df turn "0" to "FALSE" and "1" to "TRUE"
    combined_df['Important Information'] = combined_df['Important Information'].replace({0: 'FALSE', 1: 'TRUE'})

    #rename the dataframe to dataframe1
    dataframe1 = combined_df
    
    dataframe1.to_excel('dataframe1.xlsx', index = False)
    
    return dataframe1

def combine_booking_sheets():
    # File path to your Excel file
    file_path = 'Booking Stats.xlsx'  # Replace with your actual file path

    # List of month names to include
    month_names = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]

    # Load Excel file
    xlsx = pd.ExcelFile(file_path)

    # Initialize an empty DataFrame to store combined data
    combined_df = pd.DataFrame()

    # Iterate through each sheet
    for sheet_name in xlsx.sheet_names:
        if sheet_name in month_names:  # Only combine if the sheet is a month
            # Read the sheet into a DataFrame, using the first row as the header
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Add a column to indicate the source sheet
            df['Source Sheet'] = sheet_name
            
            # Append this DataFrame to the combined DataFrame
            combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    # Rename the combined DataFrame to dataframe2
    dataframe2 = combined_df
    
    dataframe2 = manipulate_dataframe2(dataframe2)
        
    return dataframe2

def manipulate_dataframe2(dataframe2):
    # Language codes mapping based on the provided information
    language_codes = {
        'Greek': 'GR', 'English': 'EN', 'Chinese': 'CH', 'Italian': 'IT',
        'German': 'DE', 'French': 'FR', 'Russian': 'RU', 'Spanish': 'ES',
        'Romanian': 'RO', 'Serbian': 'SR', 'Turkish': 'TR', 'Hebrew': 'HE',
        'Czech': 'CS', 'Hungarian': 'HU', 'Polish': 'PL', 'Bosnian': 'BS',
        'Albanian': 'SQ', 'Irish': 'GA', 'Norwegian': 'NO', 'Portuguese': 'PT',
        'Korean': 'KO', 'Japanese': 'JA'
    }

    # Function to map full language name to language code
    def map_language_to_code(language):
        return language_codes.get(language, None)

    # Apply the function to the 'Language' column to create a new 'Language Code' column
    # Replace 'LanguageColumnName' with the actual name of the column in dataframe2 that contains language names
    dataframe2['Language Code'] = dataframe2['language'].apply(map_language_to_code)

    return dataframe2

def add_ticket_cost():
    
    # Load the "ticket cost" sheet into a DataFrame
    ticket_cost_df = pd.read_excel('Booking Stats.xlsx', sheet_name='Ticket Cost')
    
    # Set the product code as the index for easier access
    ticket_cost_df.set_index('Product Code', inplace=True)
        
    # Concatenate the product_code and Language Code
    dataframe2['Full Product Code'] = dataframe2['product_code'] + dataframe2['Language Code']
    
    # Iterate over each row in dataframe2
    for index, row in dataframe2.iterrows():
        # Get the full product code for the current row
        full_product_code = row['Full Product Code']
        # Get the month for the current row, assuming the 'month' column format is 'Month YYYY'
        month = row['month'].split()[0]  # Take only the first part, which is the month name
        
        # Find the price in the ticket cost dataframe
        if full_product_code in ticket_cost_df.index:
            # Extract the price for the corresponding month
            price = ticket_cost_df.at[full_product_code, month]
            # Check if the price is a Series or a single value
            if isinstance(price, pd.Series):
                price = price.iloc[0]  # Take the first element of the Series
            elif isinstance(price, str):
                # Clean up the price to be a float (remove '€' and convert to float)
                price = float(price.replace('€', ''))
        else:
            # If the product code is not found, set the price to None or a default value
            price = 0
        
        dataframe2.at[index, 'Ticket Price'] = price
    
    #drop the 'Full Product Code' column
    dataframe2.drop(columns=['Full Product Code'], inplace=True)
    
    #fill the NaN values with 0
    dataframe2['Ticket Price'].fillna(0, inplace=True)
    
    dataframe2.to_excel('dataframe2.xlsx', index = False)
    
    return dataframe2

def create_successful():
    #create a new dataframe that contains only the listings with a rating of 4 or higher

    #first create  a copy of the original dataframe to operate upon
    successful = dataframe1.copy()
    
    #filter the successful visits
    if 'Overall Experience' in successful.columns:
        successful = successful[successful['Overall Experience'].isin(['Excellent(5 stars)', 'Positive (4 stars)', 'Excellent (5*)', 'Positive (4*)', '5*', '4*'])]
    
    successful.to_excel(output_loc + 'Successful.xlsx', index = False)
    
    return successful

def go_together():

    #find which tours go together. to do that we will use the groupby operator on the second dataframe
    grouped = dataframe2.groupby('seller_name',)['product_code'].agg(list).reset_index()
    grouped.columns = ['seller name', 'product code']
    grouped.to_excel(output_loc + 'grouped.xlsx', index = False)

    #now create a dictionary, so that for each seller we have the name of the tours he provides, instead of their codes
    map_together(grouped)


def map_together(grouped):

    #create a copy of the dataframe to operate upon
    grouped2 = grouped.copy()

    #map the columns to product names using the product dictionary
    grouped2['product_name'] = grouped2['product code'].apply(lambda codes: ', '.join(product_dict.get(code, '') for code in codes))
    grouped2.to_excel(output_loc + 'sellernames_tours.xlsx', index = False)



def create_dictionary():
    
    #map product codes to product titles
    mapping = dict(zip(dataframe2['product_code'] , dataframe2['product_title']))

    #Convert the dictionary to a DataFrame in order to save it as an excel file
    df_product = pd.DataFrame(list(mapping.items()), columns=['product_code', 'product_title'])

    #save the DataFrame to an Excel file
    df_product.to_excel(output_loc + 'dictionary.xlsx', index=False)

    return mapping

def recommended_stories():
    
    #We take the name of the tours of the successful visits
    recommended = successful.copy()
    
    #From the column 'Name of Product Reviewed' we seperate with the '|' and take the second part
    recommended['Tour_Name'] = recommended['Name of Product Reviewed'].str.split('|').str[1]
    
    recommended[['Tour_Name']].to_excel(output_loc + 'recommended.xlsx', index=False)

#  START
#from here and on our program starts
output_loc = './outputfiles/'

dataframe1, dataframe2 = read_files()

dataframe2 = add_ticket_cost()

#1. What does a successful tour look like?
successful = create_successful()

#create a dictionary that maps product codes to product titles
product_dict = create_dictionary()


#some breakpoints
print('dataframe1 size is ', dataframe1.size)
print('successful is' , successful.size)
print('dataframe 2 size is', dataframe2.size)

#which tours go together
go_together()

#which stories would we recommend
recommended_stories()