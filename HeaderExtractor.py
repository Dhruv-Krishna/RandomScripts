from TableExtractor import *
from nltk.tag import StanfordNERTagger
from collections import Counter

class HeaderExtractor:

    #this method will check if the headers are present in trainset and returns the count of found headers from each sheet
    def check_in_traindict(self,row):
        train_dict = {'LocationID':['Loc No.',	'Loc',	'Loc #','Fireman\'s Fund Loc #','BJI Loc. Code','LocationID'],
                      'BuildingName/No#':['Street No.','Bldg #','Bldg',	'Property Name / Owner Name','Bldg','* Bldg No.','*# of Bldgs',	'Entity','Organization','Clients',	'DealerDBA',	'DealerLegalName',	'EQZone',	'Loc.Name',	'LocationName',	'Location/Exposure',	'Location/VendorName',	'Name',	'NameInsured',	'Nameof',	'NamedInsured',	'Occupant',	'Province',	'StorageCompany',	'Store','BuildingName/No#'],
                      'StreetAddress':['Street Name',	'Street Address',	'Location Address',	'Street Address',	'STREET ADDRESS',	'Street',	'Location Address',	'*Street Address',	'Location',	'Address',	'address',	'AddressDescription',	'AddressLine1',	'Address1',	'RiskLocation',	'StreetAddress'],
                      'City':[	'City',	'Name',	'CITY',	'City or Province',	'*City',	'CityAddress',	'ListofLocation',	'CityName'],
                      'State':['State',	'ST',	'State or Country',	'*State Code',	'State',	'StateCode'],
                      'County':	['County'],
                      'Postcode/ZIP':['Zip',	'Zip Code',	'ZIP',	'*Zip',	'ZIP',	'ZipCode',	'PostalCode','Postcode/ZIP'],
                      'Country':['Country',	'CountryName'],
                      'Occupancy/Use':['Occupancy',	'Primary Occupancy',	'Location Name',	'*Occupancy',	'*OccupanyDescription',	'PropertyType',	'BuildingDescription',	'Buildings',	'Description',	'Detail',	'Function',	'LocationType',	'Occup.',	'OccupancyType',	'RiskCategory','Occupancy/Use'],
                      'Description':[	'Description',	'DESCRIPTION',	'Facility Division or Description',	'Description of location',	'*Property Type'],
                      'SOVCurrency':[	'SOVCurrency','2017 Projection Average Stock (MSRP Adjusted)',	'Stock Avg',	'Avg Stock Value',	'Stock Avg',	'Avg Stock',	'Average Value',	'Ave Stock'],
                      'MaxStockValue':['2017 Projection Maximum Stock (MSRP Adjusted)',	'Stock Peak',	'Max Stock Value',	'Stock Peak',	'Fixed and Adj Stock',	'Max Stock',	'Max Storage Values',	'Stock ($)'],
                      'SquareFootageFt':['Square Feet',	'Sq Ft',	'Area (Sq Ft)',	'Area/ Sq Feet',	'Sq. Ft.',	'Square Footage',	'*Square Footage',	'Total Sq. Ft.',	'Area/SqFt',	'CoveredSquareMeters',	'SqFt',	'Sq.Ft',	'Sq.Ft.',	'SquareFeet',	'SquareFootage',	'TotalSq.Ft.','SquareFootageFt'],
                      'NoOfStoreys':['# Floors',	'# of Stories',	'No. Stories',	'# Stories',	'*# of Stories','NoOfStoreys'],
                      'ConstructionType':['Construction',	'Const',	'Const.',	'Construction',	'Construction Description','ConstructionType'],
                      'YearBuilt':['Year Built',	'Year',	'*Orig Year Built',	'*OrigBuilt',	'*RoofCoveringLastFullyReplaced',	'*Yr.RoofCoveringLastRepl',	'AgeOfConstruction',	'Built',	'Heating',	'HVAC',	'OriginalBuilt',	'Plumbing',	'Rennovations',	'Roof',	'Updates',	'Wiring ',	'ElectricityUpdated',	'Renovated',	'RoofUpdated',	'Upgrade',	'YrBldgUpgraded',	'YrBuilt',	'Yr.Blt'],
                      'Sprinkler(Y/N)':['Sprinkler',	'Sprinklered (Y/N)',	'Sprinklers',	'Sprkl',	'Spr',	'Sprinkler (Y/N)',	'Sprinklers (Y or N)'],
                      'TheftAlarm(Y/N)':['Alarm',	'Alarm (Y/N)',	'Alarm (Y or N)','TheftAlarm(Y/N)'],
                      'FireAlarm(Y/N)':['Fire Alarm',	'Fire Protection',	'',	'FloodZone',	'Flood Zone',	'Flood Zone (MSW)','FireAlarm(Y/N)'],
                      'BuildingValues':['Building',	'Building Limit',	'Building****',	'Building Values',	'2017 Buildings'],
                      'BusinessInterruption(Value)':['BI/EE',	'Business Income and Extra Expense Limit',	'Business Income',	'100% Business Income',	'2018 Business Income & Extra Expense',	'Business Interruption'],
                      'TotalInsuredValue':['TOTAL',	'Total Insured Value',	'Total',	'Total Property',	'Total Building Value Incl Carports/ Garages/ Fences',	'2018 TIV',	'*Total TIV'],
                      'ContentsValues':['Contractors Equipment',	'Equip Breakdown (B&M)',	'Plant & Equipment',	'Contents Value',	'Ground UpInventory',	'2018 Furniture, Fixtures & Equipment',	'2018 Computer Equipment']
                      }

        dict_values = []
        for j in train_dict.values():
            for k in j:
                dict_values.append(k.lower().replace(" ",""))
        dict_values = sorted(dict_values)

        count = 0
        for h in row:
            for val in dict_values:
                if(h in val):
                    count += 1
                    break
        return count

    #cleans the header by removing None and stripping the spaces and converting to lower
    def validate_header(self,header):
        row = []
        for r in header:
            if r is None:
                row.append('')
            else:
                row.append(str(r).lower().replace(" ",""))
        #row = [r.replace(' ', '') for r in row]
        #print(row)
        no_of_cells_found_trainset = self.check_in_traindict(row)
        return no_of_cells_found_trainset

    #from the table values passed as arg, we will take first row as header and validate it based on train data
    #this method will return the count of no. of cells who header values matches with the one in trainset for each header in sheets
    def find_header(self,table):
        count = []
        #print(table)
        #print(table[0])
        valid_cellcount_in_row = (self.validate_header(table[0]))
        count.append(valid_cellcount_in_row)
        # for i in range(0,len(table)):
        #     row = table[i]
        #     #row_val = [str(cell).strip() for cell in row]
        #     ##No of distinct cells in a row
        #     #if (len(row) > 0 and len(row) - len(set(list(row_val))) <= 5):
        #     print(row)
        #     valid_cellcount_in_row = (self.validate_header(row))
        #     count.append(valid_cellcount_in_row)
        #     break
        #print(count)
        if(len(count)>1):
            return max(count)
        elif(len(count)==1):
            return count[0]
        else:
            return 0