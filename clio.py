#import what you have to
import pandas as pd
import matplotlib as plt

#display all columns
pd.set_option('display.max_columns', None)

#SET THE METHODS WE WILL BE USING

#read and instantiate dataframe
def create_dataframes():

    #read the whole excel file and then save the name of each sheet
    all = pd.ExcelFile('reviews data.xlsx')
    names = all.sheet_names

    #merge the sheet names in a new column
    dataframe1 = pd.concat([pd.read_excel(all, sheet_name = s, header= 1).assign(sheet_name = s) for s in names])

    #do the same for the second dataframe
    all = pd.ExcelFile('Booking Stats.xlsx')
    names = all.sheet_names

    #merge the sheet names in a new column
    dataframe2 = pd.concat([pd.read_excel(all, sheet_name = s).assign(sheet_name = s) for s in names])

    return dataframe1, dataframe2

#create a new dataframe that contains only the listings with a rating of 4 or higher
def create_successful():
    
    #first create  a copy of the original dataframe to operate upon
    successful = dataframe1.copy()
    
    #filter the successful visits
    if 'Overall Experience' in successful.columns:
        successful = successful[successful['Overall Experience'].isin(['Excellent(5 stars)', 'Positive (4 stars)', 'Excellent (5*)', 'Positive (4*)', '5*', '4*'])]
    
    return successful

def save_toExcel():
    successful.to_excel('Successful.xlsx', index = False)


def go_together():

    #find which tours go together. to do that we will use the groupby operator on the second dataframe
    grouped = dataframe2.groupby('seller_name',)['product_code'].agg(list).reset_index()
    grouped.columns = ['seller name', 'product code']
    grouped.to_excel('grouped.xlsx', index = False)

    #now create a dictionary, so that for each seller we have the name of the tours he provides, instead of their codes
    map_together(grouped)


def map_together(grouped):

    #create a copy of the dataframe to operate upon
    grouped2 = grouped.copy()

    #map the columns to product names using the product dictionary
    grouped2['product_name'] = grouped2['product code'].apply(lambda codes: ', '.join(product_dict.get(code, '') for code in codes))
    grouped2.to_excel('sellernames_tours.xlsx', index = False)



def create_dictionary():
    
    #map product codes to product titles
    mapping = dict(zip(dataframe2['product_code'] , dataframe2['product_title']))

    #Convert the dictionary to a DataFrame in order to save it as an excel file
    df_product = pd.DataFrame(list(mapping.items()), columns=['product_code', 'product_title'])

    #save the DataFrame to an Excel file
    df_product.to_excel('dictionary.xlsx', index=False)

    return mapping

def recommended_stories():
    
    #We take the name of the tours of the successful visits
    recommended = successful.copy()
    
    #From the column 'Name of Product Reviewed' we seperate with the '|' and take the second part
    recommended['Tour_Name'] = recommended['Name of Product Reviewed'].str.split('|').str[1]
    
    recommended[['Tour_Name']].to_excel('recommended.xlsx', index=False)

#  START
#from here and on our program starts
dataframe1, dataframe2 = create_dataframes()

#what does a successful tour look like
#now we have to create a new dataframe for the listings that have a rating of 4 and higher
successful = create_successful()

#create a dictionary that maps product codes to product titles
product_dict = create_dictionary()


#some breakpoints
print('dataframe1 size is ;', dataframe1.size)
print('successful is' , successful.size)
print('dataframe 2 size is', dataframe2.size)


#save the successful visits to an external excel file
save_toExcel()


#which tours go together
go_together()

#which stories would we recommend
recommended_stories()
