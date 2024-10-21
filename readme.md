This library can be used to consolidate the xl's from multiple sheets into single consolidated file.\
These are the parameter that needs to be passed for the functions.

Prerequisites Libraries:
    pandas
    xlsxwriter
    openpyxl
    configparser

There are some basic rules need to be followed while configuring the property file.\
    1 Important Points to be considered before execution.\
    2 Please provide the path eg: D: /folder/excel path.\
    3 Default HeaderColumnvaluesIndex will be considered as '0' if not provided.\
    4 Please provide the Column name of index worksheet which contains all the worksheets name else default value for column will be       considered as WorkSheetNames.\
    5 Recommended: Please provide column name in IndexSheet which stores Column header for consolidation. Else, please note the            column headers will be picked from one of the sheets randomly.\
    6 Default HeaderForConsolworksheet will be considered as '0' if not provided.\
    7 TransPoseConsolidationRow to be provided if any transpose consolidation needs to be done example 2,0,1 (2 represents number of       rows to limit , 0 represents first column ,1 represents 1st column for transpose).\
    8 Default ExcelOutPath will be considered as ConSolidated.xlsx.

Example:
    from Consoldation_utility import consolidateData\
    consolidateData("/Users/username/Downloads/Xlsx_testing/test.xlsx",
                "Sheet 3",2,"Names","",3,"2,0,1","/Users/username/Downloads/Xlsx_testing/consolidated_testing.xlsx")
\
    Sheet1:\
        ![img_1.png](
https://github.com/suprithvasista/ConsolidationLibrary/blob/main/img_1.png
        )
\
    Sheet2:\
        ![img_2.png](
        https://github.com/suprithvasista/ConsolidationLibrary/blob/main/img_2.png
        )
\
    Output:\
        Consolidate sheet:
\
        ![img_3.png](
        https://github.com/suprithvasista/ConsolidationLibrary/blob/main/img_3.png
        )
\
        Transposed sheet:
\
        ![img_4.png](
        https://github.com/suprithvasista/ConsolidationLibrary/blob/main/img_4.png
        )\
\
In next release will try to incorporate some database connection to load the consolidated data into DB.
