import re
from HeaderExtractor import *

class TableExtractor:
    #this will extract entire content in excel file
    def extract_full_table(self,sheet):
        max_row = sheet.max_row
        max_column = sheet.max_column
        table_values = []
        for i in range(1, max_row + 1):
            row_values = []
            for j in range(1, max_column + 1):
                cell_obj = sheet.cell(row=i, column=j)
                row_values.append(cell_obj.value)
            table_values.append(row_values)
        #print(table_values)
        return table_values

    #this will extract only the reuired table from entire table
    #header(table start) is the row having maximum no. of distinct cells
    #footer(table end) is end of sheet) - need to modify this
    def extract_exact_table(self,table):
        row_lens = []
        for r in table:
            row_lens.append(len(set(r)))
        #print(row_lens)
        exact_table = []

        req_index = row_lens.index(max(row_lens))
        for i in range(req_index,len(table)):
            table[i] = ['None' if v is None else v for v in table[i]]
            exact_table.append(table[i])

        return exact_table

    #In case of multiple sheets to identify which sheet we require for sov extraction
    def get_required_sheet(self,sheetnames,workbook):
        count = []
        for sht in sheetnames:
            #print(sht)
            sheetdata = workbook[sht]
            table = self.extract_full_table(sheetdata)
            exact_table = self.extract_exact_table(table)
            valid_cell_count_in_header = HeaderExtractor().find_header(exact_table)
            count.append(valid_cell_count_in_header)

        print(count,sheetnames)
        if(len(count)>1):
            #print(count.index(max(count)),sheetnames[count.index(max(count))])
            return sheetnames[count.index(max(count))],exact_table


    #this method will handle merge cells
    def merge_cell_handler(self,sheet):
        merge_cells = [tuple(str(i).split(':')) for i in sheet.merged_cells]
        merge_flag = 0
        if len(merge_cells) > 0:
            merge_flag = 1
        for start, end in merge_cells:
            start_col = re.findall('[A-Z]+', start)[0]
            end_col = re.findall('[A-Z]+', end)[0]
            start_row = re.findall('[0-9]+', start)[0]
            end_row = re.findall('[0-9]+', end)[0]

            alpha = [chr(i) for i in range(65, 91)] + [''.join([chr(i), chr(j)]) for i in range(65, 91)
                                                       for j in range(65, 91)]

            if (start_col == end_col) and (start_row != end_row):
                for i in range(int(start_row), int(end_row) + 1):
                    sheet[start_col + str(i)].value = sheet[start].value

            elif (start_col != end_col) and (start_row == end_row):
                for i in alpha[alpha.index(start_col):alpha.index(end_col) + 1]:
                    sheet[i + start_row].value = sheet[start].value

        return sheet, merge_flag