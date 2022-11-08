import pandas as pd
import os 

## to run the script on mac open terminal and type: python3 /Users/{USER}/desktop/{FOLDER}/intake-creator.py

# define the folder where the files currently live and where they should be exported
data_location ='/Users/jc/desktop/citationsubmissions/files/'

# define the folder where the new files should be exported to
new_file_directory = '/Users/jc/desktop/citationsubmissions/Processed Files/'

# create a list of the files within data_location
file_list = []
for file in os.listdir(data_location): 
    file_list.append(file)

# create a new list of each file within file_list that is a .xlsx file
new_list = [x for x in file_list if ".xlsx" in x]

# loop over the list of csv files
for f in new_list:
    
    # read the csv file and delete the cover page
    sheet1 = pd.read_excel(data_location + f, sheet_name='Cover')
    del sheet1
    sheet2 = pd.read_excel(data_location + f, sheet_name='ABIS (per location)')

    # Remove rows where all values are missing
    sheet2.dropna(inplace = True, how='all')

    # Create a list of the values to keep
    keep_rows = ['Business Name',
        'Address Line 1',
        'Address Line 2',
        'City',
        'State/Province',
        'ZIP Code',
        'Business Primary Phone Number',
        'Website URL',
        'Display URL',
        'Location URL',
        'Google Maps Local URL',
        'Facebook Page URL',
        'Yelp Listing URL',
        'Primary Category',
        'Additional Categories (Sushi Restaurant, Cafe, etc.)',
        'Hours of Operation',
        'Year Established',
        'Payment Types Accepted',
        'Business Slogan or Tagline',
        'Areas You Serve',
        'Neighborhood',
        'Handicap Accessible',
        'LGBTQ Friendly',
        'Accepts Reservations',
        'Keyword',
        'Business Description',
        'Company Logo:',
        'Image',
        'Image caption',
        ]

    # Drop the values that are ~NOT~ the values to keep
    sheet2 = sheet2[~sheet2['Location ABIS (Approved Business Information Sheet)'].str.contains('|'.join(keep_rows))==False]
    sheet2 = sheet2[~sheet2['Location ABIS (Approved Business Information Sheet)'].str.contains('Service Category')==True]
    sheet2 = sheet2[~sheet2['Location ABIS (Approved Business Information Sheet)'].str.contains('LEGALLY REGISTERED')==True] 

    # Define Business Name and City as python variables
    BusinessName = sheet2[sheet2['Location ABIS (Approved Business Information Sheet)'].str.contains('Business Name')]#.to_string(index=False, header=False)
    City = sheet2[sheet2['Location ABIS (Approved Business Information Sheet)'].str.contains('City')]#.to_string(index=False, header=False)

    ## Creating each filename
    # Get JUST the defined variables for Business Name (bn) and City (c)
    bn = sheet2.set_index('Location ABIS (Approved Business Information Sheet)')['Unnamed: 1']
    c = sheet2.set_index('Location ABIS (Approved Business Information Sheet)')['Unnamed: 1']

    # Define the above targeted data point as a new variable
    BusinessName1 = bn['Business Name']
    City1 = c['City']

    # print the location and filename
    print('Location:', f)
    print('File Name:', f.split("\\")[-1])

    # Export data to new sheet without Index or Headings
    sheet2.to_excel(new_file_directory + str(BusinessName1) + '- ' + str(City1) + ' - Intake' +  '.xlsx', sheet_name='Intake Form', index=False, header=False)
    print('>> Export Successful')
    print('------------')
successful = """
------------------------
Intake Export Successful
------------------------
"""
print(successful)

