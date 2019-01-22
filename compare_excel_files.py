import pandas as pd
import time
import os

output_excel_file = str(str(os.getenv('HOMEPATH')).replace("\\","/") + "/Downloads/output_"+str(time.time()).split('.')[0]+".xlsx")

def compare(file1_path, file2_path, output_path):
    file1 = pd.read_excel(file1_path, None)
    file2 = pd.read_excel(file2_path, None)

    # This would be a sheet name to sheet name comparison. So compare the number of sheets in both the files. If same we can go ahead and start doing a sheet to sheet comparison. Later we can provide a option of user selected sheet to sheet comparison
    if (file1.keys() == file2.keys()):
        print("Both files have same sheets")
        # Get the names of all sheets
        all_tabs = list(file1.keys())
        print(all_tabs)

        # Output excel file to write the comparison results
        output_excel_file = pd.ExcelWriter(str(output_path+'output_'+str(time.time()).split('.')[0]+'.xlsx'))

        # Now loop through all the sheets in both the files
        for sheet in all_tabs:
            print("Comparing Sheet -", sheet)

            # Scan each sheet in two seperate dataframes
            file1 = pd.read_excel(file1_path, sheetname=sheet, na_values="", dtype=str)
            file2 = pd.read_excel(file2_path, sheetname=sheet, na_values="", dtype=str)

            # Replace all NaN values with blanks
            file1.fillna("", inplace=True)
            file2.fillna("", inplace=True)

            # all_indexes = list(set(list(file1.index.values) + list(file2.index.values)))
            # rslt_df = pd.DataFrame(None,all_indexes,file1.columns,dtype=str)

            # Compare if both the sheets have same number and names of columns
            if (file1.columns.equals(file2.columns)):
                print("Both sheets have same columns")

                # Compare if both the sheets have same number of rows
                if (file1.index.equals(file2.index)):
                    print("Both sheets have same index")

                    # Create a empty dataframe to store the comparison result and also write directly to an excel file or anywhere or just show the user anyway....
                    rslt_df = pd.DataFrame(None, file1.index, file1.columns, dtype=str)

                    # Now iterate through each row for this sheet
                    for index, row in file1.iterrows():

                        # Get the same numbered rows from both the files. squeeze() is used for the other file's row to remove an additional dimension (getting data using index methods seems to return a multidimentional series) Can be checked for other options in future
                        file1_row = row
                        file2_row = file2[file2.index == index].squeeze()

                        # result seems to be a bool type dataype!! But it generates a matrix sort of structure with column names as index and TRUE/FALSE as the index value
                        result = (file1_row == file2_row)

                        # Now iterate through the result matrix to figure out which columns have the comparison result as TRUE or FALSE
                        for item in result.iteritems():
                            if (item[1] == False):
                                # print("Values mismatched in row -",index+1,"for column",item[0])
                                # print("File 1 -",file1_row[item[0]],"File 2 - ",file2_row[item[0]])
                                # If the value is false (Cell value in both the file is different), make the user aware in some way
                                error_str = file1_row[item[0]] + " -> " + file2_row[item[0]]
                                # Set the above string with values from both the files in the result dataframe.
                                rslt_df.set_value(index, item[0], error_str)
                            else:
                                # If both the file's value is same, set value from any one of the file in the result dataframe.
                                rslt_df.set_value(index, item[0], file1_row[item[0]])
                    # rslt_df.to_excel(output_excel_file,sheet_name=sheet,index=False)
                else:
                    print("IN ELSE (TO DO): If number of rows are not same")

                    # print("SUGGESTION: First column can be used as a reference for overall comparison")

                    # Again read both the sheets but with first column as index and data type as string (temporary)
                    file1 = pd.read_excel(file1_path, sheetname=sheet, na_values="", index_col=0)
                    file2 = pd.read_excel(file2_path, sheetname=sheet, na_values="", index_col=0)

                    # Remove all the null values
                    file1.fillna("", inplace=True)
                    file2.fillna("", inplace=True)

                    # Merge indexes from both the sheets. This would give us a list of all the indexes which might be missing in any one of the file
                    all_indexes = list(set(list(file1.index.values) + list(file2.index.values)))
                    # Create a new data frame to include the above indexes
                    rslt_df = pd.DataFrame(None, all_indexes, file1.columns, dtype=str)

                    # Now iterate through both the files using the combined index list
                    for index in all_indexes:

                        # Same as above, to remove additional dimension from the series
                        file1_row = file1[file1.index == index].squeeze()
                        file2_row = file2[file2.index == index].squeeze()

                        # If the index is present in both the files, just do a cell to cell comparison
                        if ((file1_row.empty == False) & (file2_row.empty == False)):
                            # result would include a True/False value for each cell of the row based on the comparison result. TO DO: Need to figure out comparison between different datatypes
                            result = (file1_row == file2_row)
                            for item in result.iteritems():
                                if (item[1] == False):
                                    # print("Values mismatched in row -",index+1,"for column",item[0])
                                    # print("File 1 -",file1_row[item[0]],"File 2 - ",file2_row[item[0]])
                                    # This string needs to be updated later to include comparison between float, int datatypes (If dtype=str is not provided while scanning the file, this line throws an error)
                                    error_str = str(file1_row[item[0]]) + " -> " + str(file2_row[item[0]])
                                    rslt_df.set_value(index, item[0], error_str)
                                else:
                                    rslt_df.set_value(index, item[0], file1_row[item[0]])

                        # If the index is not present in one of the file, include that row as it is in the result
                        elif ((file1_row.empty == True) & (file2_row.empty == False)):
                            # print("Additional row present in File 2 - ",file2_row)
                            #file2_row[0] = str("+ " + file2_row[0])
                            rslt_df.loc[index] = file2_row
                        elif ((file1_row.empty == False) & (file2_row.empty == True)):
                            # print("Additional row present in File 1 - ",file1_row)
                            #file1_row[0] = str("+ " + file1_row[0])
                            rslt_df.loc[index] = file1_row
                    # print(rslt_df)

            else:
                print("IN ELSE (TO DO): Combining columns and then comparing")

            rslt_df.to_excel(output_excel_file, sheet_name=sheet, index=False)

        output_excel_file.save()
    else:
        print("IN ELSE (TO DO): Merge these two lists")

    return True

def compare_sheets(file1_path, file2_path, sheet_name):

    print("Comparing Sheet -", sheet_name)

    # Scan each sheet in two seperate dataframes
    file1 = pd.read_excel(file1_path, sheetname=sheet_name, na_values="")
    file2 = pd.read_excel(file2_path, sheetname=sheet_name, na_values="")

    # Replace all NaN values with blanks
    file1.fillna("", inplace=True)
    file2.fillna("", inplace=True)

    # Compare if both the sheets have same number and names of columns
    if (file1.columns.equals(file2.columns)):
        print("Both sheets have same columns")

        # Compare if both the sheets have same number of rows
        if (file1.index.equals(file2.index)):
            print("Both sheets have same index")

            # Create a empty dataframe to store the comparison result and also write directly to an excel file or anywhere or just show the user anyway....
            rslt_df = pd.DataFrame(None, file1.index, file1.columns, dtype=str)

            # Now iterate through each row for this sheet
            for index, row in file1.iterrows():

                # Get the same numbered rows from both the files. squeeze() is used for the other file's row to remove an additional dimension (getting data using index methods seems to return a multidimentional series) Can be checked for other options in future
                file1_row = row
                file2_row = file2[file2.index == index].squeeze()

                # result seems to be a bool type dataype!! But it generates a matrix sort of structure with column names as index and TRUE/FALSE as the index value
                result = (file1_row == file2_row)

                # Now iterate through the result matrix to figure out which columns have the comparison result as TRUE or FALSE
                for item in result.iteritems():
                    if (item[1] == False):
                        # print("Values mismatched in row -",index+1,"for column",item[0])
                        # print("File 1 -",file1_row[item[0]],"File 2 - ",file2_row[item[0]])
                        # If the value is false (Cell value in both the file is different), make the user aware in some way
                        error_str = file1_row[item[0]] + " -> " + file2_row[item[0]]
                        # Set the above string with values from both the files in the result dataframe.
                        rslt_df.set_value(index, item[0], error_str)
                    else:
                        # If both the file's value is same, set value from any one of the file in the result dataframe.
                        rslt_df.set_value(index, item[0], file1_row[item[0]])
            # rslt_df.to_excel(output_excel_file,sheet_name=sheet,index=False)
        else:
            print("IN ELSE (TO DO): If number of rows are not same")

            # print("SUGGESTION: First column can be used as a reference for overall comparison")

            # Again read both the sheets but with first column as index and data type as string (temporary)
            file1 = pd.read_excel(file1_path, sheetname=sheet_name, na_values="", index_col=0)
            file2 = pd.read_excel(file2_path, sheetname=sheet_name, na_values="", index_col=0)

            # Remove all the null values
            file1.fillna("", inplace=True)
            file2.fillna("", inplace=True)

            # Merge indexes from both the sheets. This would give us a list of all the indexes which might be missing in any one of the file
            all_indexes = list(set(list(file1.index.values) + list(file2.index.values)))
            # Create a new data frame to include the above indexes
            rslt_df = pd.DataFrame(None, all_indexes, file1.columns, dtype=str)

            # Now iterate through both the files using the combined index list
            for index in all_indexes:

                # Same as above, to remove additional dimension from the series
                file1_row = file1[file1.index == index].squeeze()
                file2_row = file2[file2.index == index].squeeze()

                # If the index is present in both the files, just do a cell to cell comparison
                if ((file1_row.empty == False) & (file2_row.empty == False)):
                    # result would include a True/False value for each cell of the row based on the comparison result. TO DO: Need to figure out comparison between different datatypes
                    result = (file1_row == file2_row)
                    for item in result.iteritems():
                        if (item[1] == False):
                            # print("Values mismatched in row -",index+1,"for column",item[0])
                            # print("File 1 -",file1_row[item[0]],"File 2 - ",file2_row[item[0]])
                            # This string needs to be updated later to include comparison between float, int datatypes (If dtype=str is not provided while scanning the file, this line throws an error)
                            error_str = str(file1_row[item[0]]) + " -> " + str(file2_row[item[0]])
                            rslt_df.set_value(index, item[0], error_str)
                        else:
                            rslt_df.set_value(index, item[0], file1_row[item[0]])

                # If the index is not present in one of the file, include that row as it is in the result
                elif ((file1_row.empty == True) & (file2_row.empty == False)):
                    # print("Additional row present in File 2 - ",file2_row)
                    file2_row[0] = str("+ " + file2_row[0])
                    rslt_df.loc[index] = file2_row
                elif ((file1_row.empty == False) & (file2_row.empty == True)):
                    # print("Additional row present in File 1 - ",file1_row)
                    file1_row[0] = str("+ " + file1_row[0])
                    rslt_df.loc[index] = file1_row
            # print(rslt_df)

    else:
        print("IN ELSE (TO DO): Combining columns and then comparing")
    #rslt_df.to_excel(output_excel_file, sheet_name=sheet_name, index=False)
    return rslt_df
