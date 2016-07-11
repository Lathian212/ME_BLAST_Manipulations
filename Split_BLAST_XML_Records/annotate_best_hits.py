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
# This workbook will be saved under another name with merge in filename
wb_best_hit = load_workbook(
    'best_hits_Blastn_viral_queries_against_Lh_TSA.xlsx', read_only= False)
# To perserve starting input best hit workbook save it to another file
wb_best_hit.save('merged.xlsx')
wb_merge = load_workbook('merged.xlsx')
ws_merge = wb_merge.active
merge_row = 2
merge_col = 6
# Put in header from report (just as an exercise)
for i in range(1,31):
    val = ws_report.cell(row = 1 , column = i).value
    merge_cell = ws_merge.cell(row = 2, column = (6 + i -1))
    merge_cell.value = val
# Range of query sequences in best_hits file
"""found_one_match error trapping code was run ONCE and it showed every best
    hit query had one and only one match. So this will be edited out after
    this commit.
"""

for i in range (3, 4292):
    # Flag to test that each best hit query sequence find one and only one match
    found_one_match = False
    best_hit_query = ws_merge.cell(row = i, column = 2).value
    for j in range (2, 4291):
        report_val = ws_report.cell(row = j, column = 20).value
        if best_hit_query == report_val and found_one_match == False:
            #print('For Query_' + str(i-2) + 'found a match at Report row ' +
            #str(j + 2))
            found_one_match = True
        elif best_hit_query == report_val and found_one_match == True:
            print('ERROR! For Query_' + str(i-2) + 
                   'found ANOTHER match at Report row ' + str(j + 2))
    if found_one_match == False:
        print('ERROR! For Query_' + str(i-2) + ' found no match')
"""Range of viral query sequences in report_queries is 0 to 42888. Just add
2 to get the row in the actual report    
print(report_queries[0])
print(report_queries[4288])
"""
#Now get each best hit query one at a time; look for a match in the 
# report_queries list. If so call a copy information. Have a boolean
# flag to make sure there is a match and there is no more then one match.
wb_merge.save('merged.xlsx')
print('Program done.')