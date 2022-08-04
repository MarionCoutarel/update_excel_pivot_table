# update_excel_pivot_table
Creating series of Excel files with pivot tables updating data in template Excel files and refreshing pivot table


Checking ETL output in a data migration project, I need to create 4 Excel pivot tables per society to check financial balance

- STEP 1 : 
get datas from ETL output csv file

- STEP 2 : 
Using OpenPyXl, for each society, replacing datas in pivot tables data sheet and save it under a new generated name in a new folder

- STEP 3 :
Using win32, refresh all the pivot tables in the generated files so they display the right values.
