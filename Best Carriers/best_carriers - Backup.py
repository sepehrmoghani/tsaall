import pandas as pd

#Quote Head: Date,From_location,To_Location,Service Type,Unit Type,Unit Value,Cubic Size,Weight,Pallet Size,Original Cost,Calculated Distance

def main():
    carrier_quote = pd.read_csv("C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/carriers.csv")
    data = input("Please enter the excel quote (without .xlsx): ")
    quote_quote = pd.read_excel(data + ".xlsx")

    results = adjust_table(quote_quote, carrier_quote)

    # Convert the list of matching carriers to a comma-separated string for each row
    results_str = [', '.join(map(str, res)) for res in results]
    
    # Add the results as a new column to the original DataFrame
    quote_quote['Matching Carriers'] = results_str
    
    final_quote = collect_unique_carriers(quote_quote, 'Matching Carriers') #list of all the unique carriers
    final_quote_str = ', '.join(final_quote)  # Convert the list to a comma-separated string
    
    # Save the DataFrame with the new columns back to an Excel file
    output_excel_path = data + "_with_matching_carriers.xlsx"  # Add a suffix to the original filename
    quote_quote.to_excel(output_excel_path, index=False)
    print(f"Results saved to {output_excel_path}")

    # Save the final_quote_str to a new text file
    text_file_path = data + "_list_of_carriers.txt"  # Create a new text file name based on the original Excel filename
    with open(text_file_path, 'w') as f:
        f.write(final_quote_str)
    print(f"Final quote saved to {text_file_path}")

def find_matching_carriers(quote, carrier):
    results = []

    for index, row in quote.iterrows():
        matching_carriers_list = []

        unit_service_type = row['Service Type Required']
        unit_type = row['Unit Type']
        from_area = location_finder(row['From'])[0]  # return statistic_area, state
        to_area = location_finder(row['To'])[0]
        
        # Remove 'Rest of ' prefix if present
        from_area = from_area.replace('Rest of ', '') if from_area else from_area
        to_area = to_area.replace('Rest of ', '') if to_area else to_area

        for _, carrier_row in carrier.iterrows():
            service_types = carrier_row['Service Type'].split('; ') #satchel, carton, pallet
            service_area_from = carrier_row.get('Service Area From', '').split('; ') #City or State from
            service_area_to = carrier_row.get('Service Area To', '').split('; ') #City or State to
            service_types_network = carrier_row.get('Network', '').split('; ') #local, regional, interstate
            
            if unit_type in service_types and from_area in service_area_from and to_area in service_area_to and unit_service_type in service_types_network:
                matching_carriers_list.append(carrier_row['Carrier Name'])

        results.append(matching_carriers_list)
        
    return results

def adjust_table(quote, carrier):

    capital_cities = ['Sydney', 'Brisbane', 'Melbourne', 'Adelaide', 'Hobart', 'Perth']
    quote['Service Type Required'] = quote.apply(
        lambda row: 
        'Local' if (location_finder(row['From'])[0] == location_finder(row['To'])[0] and location_finder(row['From'])[0] in capital_cities) 
        else ('Regional' if location_finder(row['From'])[1] == location_finder(row['To'])[1] 
              else 'Interstate'), 
        axis=1)

    quote.loc[((quote['Cubic Size'].isnull()) | (quote['Cubic Size'] == 0)) & (quote['Weight']/quote['Unit Value'] <= 5), 'Cubic Size'] = 0.015 #satchels are less than 5kg
    quote.loc[((quote['Cubic Size'].isnull()) | (quote['Cubic Size'] == 0)) & (quote['Weight']/quote['Unit Value'] > 5) & (quote['Weight']/quote['Unit Value'] <= 35), 'Cubic Size'] = 0.08 #Carton over 23kg are labels heavy
    quote.loc[((quote['Cubic Size'].isnull()) | (quote['Cubic Size'] == 0)) & (quote['Weight']/quote['Unit Value'] > 35), 'Cubic Size'] = 1.44 #anything over 35kg can be considered pallet

    quote.loc[(quote['Cubic Size'] <= 0.015) & (quote['Weight'] <= 5), 'Unit Type'] = 'Satchel'
    quote.loc[(quote['Cubic Size'] > 0.015) & (quote['Cubic Size'] <= 0.08) | (quote['Weight']/quote['Unit Value'] > 5) & (quote['Weight']/quote['Unit Value'] <= 35), 'Unit Type'] = 'Carton'
    quote.loc[(quote['Cubic Size'] > 0.08) & (quote['Weight']/quote['Unit Value'] > 35), 'Unit Type'] = 'Pallet'

    quote['To'] = quote['To'].astype(str)
    quote['From'] = quote['From'].astype(str)

    results = find_matching_carriers(quote, carrier)
    
    return results

# Read the CSV file once and store it in a variable
file_path = "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/suburbs.csv"
suburbs = pd.read_csv(file_path)

# Dictionary for hard-coded values
postcode_prefixes = {
    '2': ('Rest of NSW', 'NSW'),
    '3': ('Rest of VIC', 'VIC'),
    '4': ('Rest of QLD', 'QLD'),
    '5': ('Rest of SA', 'SA'),
    '6': ('Rest of WA', 'WA'),
    '7': ('Rest of TAS', 'TAS'),
    '8': ('Rest of NT', 'NT')
}

def location_finder(postcode):
    global suburbs  # Use the pre-loaded DataFrame
    if not (isinstance(postcode, int) or postcode.isnumeric()):
        return None, None
    
    postcode = int(postcode)
    matching_suburb = suburbs[suburbs['postcode'] == postcode]
    
    if not matching_suburb.empty:
        return matching_suburb.iloc[0]['statistic_area'], matching_suburb.iloc[0]['state']
    
    return postcode_prefixes.get(str(postcode)[0], (None, None))



def collect_unique_carriers(df, column_name):
    unique_carriers = []
    for carriers in df[column_name]:
        carriers_list = carriers.split(", ")
        for carrier in carriers_list:
            if carrier not in unique_carriers:
                unique_carriers.append(carrier)
    unique_carriers.sort()
    return unique_carriers

if __name__=="__main__":
    main()