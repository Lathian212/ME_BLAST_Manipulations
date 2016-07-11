'''
Created on Jul 10, 2016

@author: jonathan
The purpose of this script is to take the file:
best_hits_Blastn_viral_queries_against_Lh_TSA.xlsx
and merge in the annotations from the blastx results from the Lipkin lab in:
CII-MWP0002.contigs_singletons.blastx.viral_report.xlsx
'''
""" To work with Excel .xlsx files I am using openpyxl """
from openpyxl import Workbook
from openpyxl import load_workbook
#MAIN STARTS
#Most of the report has 30 columns and it does in the header
wb_report = load_workbook(
    'CII-MWP0002.contigs_singletons.blastx.viral_report.xlsx', read_only= False)
ws_report = wb_report.active
wb_merge = load_workbook(
    'best_hits_Blastn_viral_queries_against_Lh_TSA.xlsx', read_only= False)
ws_merge = wb_merge.active
merge_row = 2
merge_col = 6
# Put in header from report (just as an exercise)
for i in range(1,31):
    val = ws_report.cell(row = 1 , column = i).value
    merge_cell = ws_merge.cell(row = 2, column = (6 + i -1))
    merge_cell.value = val
# Scan in all query sequences in report which are in col 20, rows 2 to 4290
report_queries = []
for i in range (2, 4291):
    temp_val = ws_report.cell(row = i, column = 20).value
    report_queries.append(temp_val)
"""Range of viral query sequences in report_queries is 0 to 42888. Just add
2 to get the row in the actual report    
print(report_queries[0])
print(report_queries[4288])
"""
#Now get each best hit query one at a time; look for a match in the 
# report_queries list. If so call a copy information. Have a boolean
# flag to make sure there is a match and there is no more then one match.

#wb_merge.save('merged.xlsx')
print('Program done.')