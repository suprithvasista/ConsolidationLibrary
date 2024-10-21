import sys
import pandas as pd
import configparser
import os
from pathlib import Path
from datetime import datetime


# End time function.
def end_time():
    end = datetime.now()
    print("Script end Time : ", end)


def path_trim(excel_path):
    path = excel_path
    excel_path = path.replace("\\", "/")
    print('Excel file paths : ', excel_path)
    return excel_path


def identify_variabl(TransposeVariable):
    if isinstance(TransposeVariable, str):
        values = TransposeVariable.split(',')
        z = values[0] == '0'
        if len(values) <= 2:
            print("The Transpose values should atleast have three numbers eg:- 0,0,1 .")
            end_time()
            sys.exit(1)
        if z:
            return "No row limiter"
        else:
            return "With Limiter"


def consolidateData(excel_path, index_sheet_name, header_base_worksheets, Col_name_sheet_consolidation,
                    Column_name_Header_name_for_Consolidation, HeaderForConsolidatedworksheet, TransPoseRownumber,
                    Excel_ouput):
    def datarows(headerMethod):
        for index, row in data_frame.iterrows():
            works_sheet_name = row[Col_name_sheet_consolidation]
            work_sheet_workbook.append(works_sheet_name)
            if headerMethod == "Y":
                header_data_frame = row[Column_name_Header_name_for_Consolidation]
                header_names.append(header_data_frame)

    if excel_path == "":
        print("Excel path not passed")
        end_time()
        sys.exit(1)
    if index_sheet_name == "":
        print("Index sheet name not passed.")
        end_time()
        sys.exit(1)
    if header_base_worksheets == "":
        print("Index sheet name header not passed consider default header: 0")
        header_base_worksheets = int(0)
    if Col_name_sheet_consolidation == "":
        print("Index sheet name not passed defaulting to WorkSheetNames.")
        Col_name_sheet_consolidation = "WorkSheetNames"
    if Column_name_Header_name_for_Consolidation == "":
        print("Column name not passed which stores header names for consolidation.")
        # header_consol_name = "N"
    if HeaderForConsolidatedworksheet == "":
        print("Header for Consolidating Sheets not passed, defaulting header to : 0")
        HeaderForConsolidatedworksheet = int(0)
    if Excel_ouput == "":
        print("Output Excel path Not provided Reverting to default mode.")
        Excel_ouput = "ConSolidated.xlsx"
    if TransPoseRownumber == "":
        print("No Transpose consolidation.")

    r1 = path_trim(excel_path)
    path = Path(r1)

    try:
        if not path.is_file():
            raise FileNotFoundError("Excel file not found in : ", path)
    except FileNotFoundError as e:
        print("An error occurred: ", e)
        end_time()
        sys.exit(101)

    work_sheet_workbook = []
    header_names = []
    try:
        if Column_name_Header_name_for_Consolidation != "":
            data_frame = pd.read_excel(excel_path, sheet_name=index_sheet_name, header=int(header_base_worksheets),
                                       usecols=[Col_name_sheet_consolidation,
                                                Column_name_Header_name_for_Consolidation],
                                       na_filter=False)
            print(data_frame.head(15))
            datarows("Y")
        else:
            df = pd.read_excel(excel_path, sheet_name=index_sheet_name, header=int(header_base_worksheets),
                               usecols=[
                                   Col_name_sheet_consolidation])  # , na_filter=False)
            data_frame = df.dropna()
            datarows("N")

    except MemoryError:
        print('Please reduce the sample size or free up memory some memory before execution.')
        end_time()
        sys.exit(101)
    except KeyboardInterrupt:
        print('KeyBoard interrupt Occurred.')
        end_time()
        sys.exit(101)
    except Exception as e:
        print("An error occurred: ", e, ".")
        if type(e).__name__ == "ValueError":
            e_val = str(e)
            if "Usecols" in e_val:
                print("Please Provide the Column name of INDEX worksheet, consisting all "
                      "the worsheet names in the workbook or rename the the column as 'WorkSheetNames'")
                end_time()
                sys.exit(102)
            elif "invalid literal" in e_val:
                print("Header Value should be integer/number eg: 0,1,2")
                end_time()
                sys.exit(103)
            elif "Worksheet named" in e_val:
                print("Please configure the right index sheet name in config file.")
                end_time()
                sys.exit(104)
            else:
                end_time()
                sys.exit(105)
    work_sheet_workbook = [item for item in work_sheet_workbook if item != '']
    header_names = [head_item for head_item in header_names if head_item != '']
    if len(work_sheet_workbook) > 0:
        print(f'The worksheet name are {work_sheet_workbook} .')
    else:
        print('No Worksheet names are present in the index Worksheet')
        end_time()
        sys.exit(105)

    try:
        xls = pd.ExcelFile(excel_path)
        sheet_names_exl_method = xls.sheet_names
        sheet_names_exl_method_dict = {name: True for name in sheet_names_exl_method}
        print('List from Xl method: ', sheet_names_exl_method)
    except MemoryError:
        print('Out of memory while assigning sheet names to DICT.')
        end_time()
        sys.exit(101)
    except Exception as e:
        print('Error occurred :', str(e))
        end_time()
        sys.exit(1)

    mis_matched_worksheet_present_index_not_work_sheet = []
    mis_matched_worksheet_not_index = []
    for sheet_nam1 in work_sheet_workbook:
        if sheet_nam1 not in sheet_names_exl_method_dict:
            mis_matched_worksheet_present_index_not_work_sheet.append(sheet_nam1)
            # mis_matched_worksheet.append(('WorksheetXLMethod',sheet_nam1))
    for sheet_nam2 in sheet_names_exl_method:
        if sheet_nam2 not in work_sheet_workbook:
            mis_matched_worksheet_not_index.append(sheet_nam2)

    if len(mis_matched_worksheet_present_index_not_work_sheet) == 0:
        print(f'Worksheets are matched and are correctly mapped in {index_sheet_name} Sheet.')
    else:
        print(
            f"Missing files from worksheet present in index :{mis_matched_worksheet_present_index_not_work_sheet}.\n"
            f"Please rectify this error bring Integrity in Both {index_sheet_name} sheet names and actual worksheets names.")

    final_list = [value for value in work_sheet_workbook if
                  value not in mis_matched_worksheet_present_index_not_work_sheet]

    if Column_name_Header_name_for_Consolidation == "" or len(header_names) == 0:
        df_columns_header = pd.read_excel(excel_path, sheet_name=final_list[0],
                                          header=int(HeaderForConsolidatedworksheet),
                                          na_filter=False, nrows=5)
        header_names.extend(df_columns_header.columns)
    header_names = [columns.upper() for columns in header_names]
    print('Header for consolidation are : ', header_names)

    if TransPoseRownumber != "":
        Conditioner = identify_variabl(TransPoseRownumber)
    else:
        Conditioner = "No Process"

    try:
        Excel_ouput = path_trim(Excel_ouput)
        if os.path.exists(Excel_ouput):
            os.remove(Excel_ouput)
        if os.path.exists(Excel_ouput):
            writer = pd.ExcelWriter(Excel_ouput, engine='openpyxl', mode='a')
        else:
            writer = pd.ExcelWriter(Excel_ouput, engine='xlsxwriter')
        Skipper = 0
        Summary_skipper = 0
        Conditioner_skipper = 0
        for sheet_final in final_list:
            consolidated_df = pd.read_excel(excel_path, sheet_name=sheet_final,
                                            header=int(HeaderForConsolidatedworksheet),
                                            na_filter=False)
            size_chunk = len(consolidated_df)
            if all(col in consolidated_df.columns.str.upper() for col in header_names):
                print('Generating files in : ', Excel_ouput)
                consolidated_df.columns = consolidated_df.columns.str.upper()
                upp_header = [val.upper() for val in header_names]
                consolidated_df = consolidated_df[upp_header]
                if Skipper == 0:
                    consolidated_df.to_excel(writer, sheet_name='Consolidated', startrow=Skipper, index=False)
                else:
                    consolidated_df.to_excel(writer, sheet_name='Consolidated', header=False, startrow=Skipper + 1,
                                             index=False)
                Skipper = Skipper + size_chunk
            else:
                summary_data = {'Column_Name': [sheet_final],
                                'Description': 'Does not have matching headers sheet skipped'}
                df = pd.DataFrame(summary_data)
                if Summary_skipper == 0:
                    df.to_excel(writer, sheet_name='Summary', startrow=Summary_skipper, index=False)
                else:
                    df.to_excel(writer, sheet_name='Summary', startrow=Summary_skipper, header=False, index=False)
                Summary_skipper = Summary_skipper + 1
            if Conditioner == "With Limiter":
                values = TransPoseRownumber.split(',')
                values = [int(val) for val in values]
                row_limiter = values[0]
                Column_header_transpose = values[1]
                Column_consider_without_row_limiter = values[1:]
                df_transeposer = pd.read_excel(excel_path, sheet_name=sheet_final, nrows=int(row_limiter),
                                               usecols=Column_consider_without_row_limiter, na_filter=False,
                                               index_col=[Column_header_transpose])
                df_transeposer_col_name = pd.read_excel(excel_path, sheet_name=sheet_final, nrows=int(row_limiter),
                                                        usecols=[Column_header_transpose], na_filter=False)
                df_transeposer_col_name_fin = df_transeposer_col_name.columns[0]
                df_transeposer = df_transeposer.T.reset_index().rename(columns={'index': df_transeposer_col_name_fin})
                df_t_len = len(df_transeposer)
                if Conditioner_skipper == 0:
                    df_transeposer.to_excel(writer, sheet_name='TransPoserSheet', startrow=Conditioner_skipper,
                                            index=False)
                else:
                    df_transeposer.to_excel(writer, sheet_name='TransPoserSheet', startrow=Conditioner_skipper + 1,
                                            header=False, index=False)
                Conditioner_skipper = Conditioner_skipper + df_t_len
            elif Conditioner == "No row limiter":
                values = TransPoseRownumber.split(',')
                values = [int(val) for val in values]
                Column_header_transpose = values[1]
                Column_consider_without_row_limiter = values[1:]
                df_transeposer = pd.read_excel(excel_path, sheet_name=sheet_final,
                                               usecols=Column_consider_without_row_limiter, na_filter=False,
                                               index_col=[Column_header_transpose])
                df_transeposer_col_name = pd.read_excel(excel_path, sheet_name=sheet_final, nrows=int(row_limiter),
                                                        usecols=[Column_header_transpose], na_filter=False)
                df_transeposer_col_name_fin = df_transeposer_col_name.columns[0]
                df_transeposer = df_transeposer.T.reset_index().rename(columns={'index': df_transeposer_col_name_fin})
                df_t_len = len(df_transeposer)
                if Conditioner_skipper == 0:
                    df_transeposer.to_excel(writer, sheet_name='TransPoserSheet', startrow=Conditioner_skipper,
                                            index=False)
                else:
                    df_transeposer.to_excel(writer, sheet_name='TransPoserSheet', startrow=Conditioner_skipper + 1,
                                            header=False, index=False)
                Conditioner_skipper = Conditioner_skipper + df_t_len
                # changes to nullify future warnings for save
        writer.close()
    except MemoryError:
        print('Please reduce the sample size or free up memory some memory before execution.')
        end_time()
        sys.exit(101)
    except KeyboardInterrupt:
        print('KeyBoard interrupt Occurred.')
        end_time()
        sys.exit(101)
    except Exception as e:
        print("An error occurred: ", e, ".")
        e_val = str(e)
        if "Worksheet named" in e_val:
            print("Please verify if all worksheets are available which are mentioned in Index Worksheet")
            end_time()
            sys.exit(104)
        elif "invalid literal" in e_val:
            print("Header Value for Consolidated sheets should be integer/number eg: 0,1,2")
            end_time()
            sys.exit(103)
        else:
            end_time()
            sys.exit(1)

    end_time()
