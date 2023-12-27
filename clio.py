#import what you have to
import pandas as pd
import matplotlib as plt
import os

#SET THE METHODS WE WILL BE USING

output_loc = './outputfiles/'

#read and instantiate dataframe
def combine_review_sheets():
    # File path to your Excel file
    file_path = 'reviews data.xlsx'  # Replace with your actual file path

    # Load Excel file
    xlsx = pd.ExcelFile(file_path)

    # Collect all unique column titles from each sheet
    all_columns = []

    # Iterate through each sheet to collect column names
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1, nrows=0)  # Read only the second row for column names
        all_columns.extend([col for col in df.columns if col not in all_columns and not 'Unnamed' in str(col)])

    # Initialize an empty DataFrame to store combined data
    combined_df = pd.DataFrame()

    # Iterate through each sheet
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
    
    return dataframe1

def combine_booking_sheets():
        # File path to your Excel file
    file_path = 'Booking Stats.xlsx'  # Replace with your actual file path

    # Load Excel file
    xlsx = pd.ExcelFile(file_path)

    # Initialize an empty DataFrame to store combined data
    combined_df = pd.DataFrame()

    # Iterate through each sheet
    for sheet_name in xlsx.sheet_names:
        # Read the sheet into a DataFrame, using the first row as the header
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Add a column to indicate the source sheet
        df['Source Sheet'] = sheet_name
        
        # Append this DataFrame to the combined DataFrame
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    #rename the dataframe to dataframe2
    dataframe2 = combined_df
    
    return dataframe2



#create a new dataframe that contains only the listings with a rating of 4 or higher
def create_successful():
    
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