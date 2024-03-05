import pandas as pd
import numpy as np

def main():
    carrier_quote = pd.read_csv("C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/carriers.csv")
    data = input("Please enter the excel quote (without .xlsx): ")
    quote_quote = pd.read_excel(data + ".xlsx")

    results = adjust_table(quote_quote, carrier_quote)
    
    # Convert the list of matching carriers to a comma-separated string for each row
    results_str = [', '.join(map(str, res)) for res in results]
    
    # Add the results as a new column to the original DataFrame
    quote_quote['Matching Carriers'] = results_str
    
    final_quote = collect_unique_carriers(quote_quote, 'Matching Carriers')  # List of all the unique carriers
    
    # Add each unique carrier as a new column and populate values
    for carrier in final_quote:
        if carrier == '':
            continue  # Skip empty string
        quote_quote[carrier] = quote_quote['Matching Carriers'].apply(lambda x: 1 if carrier in x else 0)

    # Drop the 'Matching Carriers' column as it's no longer needed
    quote_quote.drop('Matching Carriers', axis=1, inplace=True)

    quote_quote['From'] = quote_quote['From'].astype(int)
    quote_quote['To'] = quote_quote['To'].astype(int)
    
    # Load the carriers (this will populate the carrier_dfs dictionary)
    load_carriers()
    
    for carrier, carrier_df in carrier_dfs.items():
        carrier_df['Postcode'] = carrier_df['Postcode'].astype(int)

    matched_zones_df = match_zones_with_carriers(quote_quote, carrier_dfs)

    # Load the rates into rate_card_dfs
    load_rates()

    # Loop through each carrier to calculate rates
    for carrier in final_quote:
        if carrier == '':
            continue  # Skip empty string

        # Fetch the rate_card DataFrame for this specific carrier from the dictionary
        rate_card_for_carrier = rate_card_dfs.get(carrier, None)

        if rate_card_for_carrier is not None:
            from_zone_col = f'From Zone [{carrier}]'
            to_zone_col = f'To Zone [{carrier}]'
            quote_quote[carrier] = quote_quote.apply(
                lambda row: calculate_rate(rate_card_for_carrier, row, from_zone_col, to_zone_col) if row[carrier] == 1 else np.nan,
                axis=1
            )

    # Filter out columns that start with 'From Zone' or 'To Zone'
    filtered_df = quote_quote.filter(regex='^(?!From Zone|To Zone).*$')

    # Save the DataFrame with the new columns back to an Excel file
    output_excel_path = data + "_Rates.xlsx"  # Add a suffix to the original filename
    filtered_df.to_excel(output_excel_path, index=False)  # Saving quote_quote DataFrame, as it now contains the rates
    print(f"Results saved to {output_excel_path}")

#=============================================================

carrier_dfs = {}
def load_carriers():
    global carrier_dfs
    
    carrier_to_path = {'Allied Overnight Express Pty Ltd': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Allied Overnight Express.csv",
                       'Aussiefast Transport Solutions General': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Aussiefast.csv",
                        'Capital_transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Capital Transport.csv",
                        'Couriers Please Carton': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Couriers Please.csv",
                        'Freight Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Flight Express.csv",
                        'Followmont Transport 002': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Followmont.csv",
                        'Followmont Transport 003': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Followmont.csv",
                        'Followmont Transport H': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Followmont.csv",
                        'GKR Transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_GKR Transport.csv",
                        'GMK Logistics': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_GMK Transport.csv",
                        'Go Logistics': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Go Logistics.csv",
                        'Hi-Trans Express':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Hi Trans Express.csv",
                        'Hi-Trans Express Pallet':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Hi Trans Express.csv",
                        'Jayde Transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Jayde.csv",
                        'Jolly and Sons': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Jolly and Sons.csv",
                        'Northline': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Northline.csv",
                        'Pobjoy Transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Pobjoy.csv",
                        'Searoad Shipping': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Searoad.csv",
                        'Startrack Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Startrack.csv",
                        'TFMXpress Pty Ltd': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_TFMXpress.csv",
                        'TNT Australia Fedex': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_TNT Australia.csv",
                        'TNT Australia H Bulk Road Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_TNT Australia.csv",
                        'TNT FedEx H Road Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_TNT Australia.csv",
                        'TOLL Intermodal & Specialised':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_TOLL.csv",
                        'TOLL Ipec':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_TOLL.csv",
                        'VELLEX': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_VELLEX.csv",
                        'Xpress Freight Management': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/ZoneDetails/ZoneDetail_Xpress Freight.csv"}
    for carrier, path in carrier_to_path.items():
        carrier_dfs[carrier] = pd.read_csv(path)

#==================================================================================

rate_card_dfs = {}
def load_rates():
    global rate_card_dfs

    rates_paths = {'Allied Overnight Express Pty Ltd': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Allied Overnight Express.csv",
                   'Aussiefast Transport Solutions General': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Aussiefast_Kettah.csv",
                   'Capital_transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Capital Transport.csv",
                   'Couriers Please Carton': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Couriers Please_Kettah.csv",
                   'Freight Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Freight Express.csv",
                   'Followmont Transport 002': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Followmont 002.csv",
                   'Followmont Transport 003': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Followmont 003.csv",
                   'Followmont Transport H': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Followmont H.csv",
                   'GKR Transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_GKR Bambach.csv",
                   'GMK Logistics': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_GMK Nolans.csv",
                   'Go Logistics': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Go Logistics.csv",
                   'Hi-Trans Express':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Hi Trans Bambach.csv",
                   'Hi-Trans Express Pallet':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Hi Trans Pallet.csv",
                   'Jayde Transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Jayde.csv",
                   'Jolly and Sons': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Jolly and Sons.csv",
                   'Northline': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Northline Kettah.csv",
                   'Pobjoy Transport': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Pobjoy.csv",
                   'Searoad Shipping': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Searoad.csv",
                   'Startrack Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Startrack TSA E.csv",
                   'TFMXpress Pty Ltd': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_TFMXpress Nolan.csv",
                   'TNT Australia Fedex': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_TNT Bombach.csv",
                   'TNT Australia H Bulk Road Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_TNT Australia  H Bulk.csv",
                   'TNT FedEx H Road Express': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_TNT Fedex H Cost.csv",
                   'TOLL Intermodal & Specialised':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Toll Intermodal & Specialised TF.csv",
                   'TOLL Ipec':"C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Toll Iped Eurotech.csv",
                   'VELLEX': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Vellex Australian Flooring.csv",
                   'Xpress Freight Management': "C:/Users/Sepehr Moghani/OneDrive - Transfreight Solutions/Documents/Python/Best Carriers/RateCards/CarrierRateCard_Xpress Freight.csv"}
    for carrier, path in rates_paths.items():
        rate_card_dfs[carrier] = pd.read_csv(path)

#========================================================================================

def calculate_rate(rate_card_df, quote_row, from_zone_col, to_zone_col):
    from_zone = str(quote_row[from_zone_col])  # Casting to string
    to_zone = str(quote_row[to_zone_col])  # Casting to string
    weight = quote_row['Weight']
    dimension = quote_row['Cubic Size']
    
    rate_card_df['ZoneFrom'] = rate_card_df['ZoneFrom'].astype(str)
    rate_card_df['ZoneTo'] = rate_card_df['ZoneTo'].astype(str)

    # Filter the rate card dataframe to only include rows that match the 'From Zone' and 'To Zone'
    filtered_df = rate_card_df[(rate_card_df['ZoneFrom'] == from_zone) & (rate_card_df['ZoneTo'] == to_zone)]
    
    if filtered_df.empty:
        return None  # No matching rate card entry found
    
    # Assuming that there's only one matching rate card entry for each (ZoneFrom, ZoneTo) pair
    rate_card_row = filtered_df.iloc[0]
    
    # Initialize rate variables
    rate = None
    basic = None
    
    # Loop through the break columns and find the appropriate rate and basic values
    for i in range(1, 11):  # Assuming up to 10 breaks
        break_col = f"Break{i}"
        rate_col = f"Rate{i}"
        basic_col = f"Basic{i}"
        
        if pd.notna(rate_card_row[break_col]):
            if i == 1 and weight <= rate_card_row[break_col]:
                rate = rate_card_row[rate_col]
                basic = rate_card_row[basic_col]
                break
            elif i > 1 and weight > rate_card_row[f"Break{i-1}"] and weight <= rate_card_row[break_col]:
                rate = rate_card_row[rate_col]
                basic = rate_card_row[basic_col]
                break
    
    if rate is None or basic is None:
        return np.nan  # No appropriate break found
    
    # Calculate the rate based on cubic rate if applicable
    cubic_rate = rate_card_row['CubicRate']
    cubic_value = dimension * cubic_rate
    
    final_value = cubic_value if cubic_value > weight else weight
    final_value *= rate
    final_value += basic
    
    # Check against minimum
    minimum = rate_card_row['Minimum']
    final_value = max(final_value, minimum)
    
    return final_value

#=======================================================================================

def match_zones_with_carriers(quote_df, carrier_dfs):
    # Loop through each column in quote_df
    for carrier in quote_df.columns:
        # Check if the column name matches a key in carrier_dfs
        if carrier in carrier_dfs:
            # Initialize new columns to store the zones for this specific carrier
            from_zone_col = f'From Zone [{carrier}]'
            to_zone_col = f'To Zone [{carrier}]'
            quote_df[from_zone_col] = None
            quote_df[to_zone_col] = None

            # Loop through each row in the column
            for index, row in quote_df.iterrows():
                if row[carrier] == 1:
                    from_postcode = row['From']
                    to_postcode = row['To']
                    
                    # Get the DataFrame for this carrier
                    carrier_df = carrier_dfs[carrier]
                    
                    # Find the zones that match the 'From' and 'To' postcodes
                    from_zone_df = carrier_df[carrier_df['Postcode'] == from_postcode]
                    to_zone_df = carrier_df[carrier_df['Postcode'] == to_postcode]
                    
                    if not from_zone_df.empty:
                        from_zone = from_zone_df['Zone'].values[0]
                    else:
                        from_zone = 'Not Found'
                    
                    if not to_zone_df.empty:
                        to_zone = to_zone_df['Zone'].values[0]
                    else:
                        to_zone = 'Not Found'
                    
                    # Update the DataFrame with the matched zones for this specific carrier
                    quote_df.at[index, from_zone_col] = from_zone
                    quote_df.at[index, to_zone_col] = to_zone

    return quote_df  # Now updated with multiple 'From Zone' and 'To Zone' columns


#===================================================================================================

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

#=========================================================================

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
    quote.loc[(quote['Cubic Size'] >= 0.015) & (quote['Cubic Size'] <= 0.08) | (quote['Weight']/quote['Unit Value'] >= 5) & (quote['Weight']/quote['Unit Value'] <= 35), 'Unit Type'] = 'Carton'
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

#==============================================================================

def location_finder(postcode):
    global suburbs  # Use the pre-loaded DataFrame
    if not (isinstance(postcode, int) or postcode.isnumeric()):
        return None, None
    
    postcode = int(postcode)
    matching_suburb = suburbs[suburbs['postcode'] == postcode]
    
    if not matching_suburb.empty:
        return matching_suburb.iloc[0]['statistic_area'], matching_suburb.iloc[0]['state']
    
    return postcode_prefixes.get(str(postcode)[0], (None, None))

#==============================================================================================

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