'''
Created on Jul 9, 2016

@author: jonathan

I realized that the task of adding the best hit for each blastn viral query
against the Leptopilina hetertoma TSA to the:
Blastn_viral_queries_againstLh_TSA.xlsx Spreadsheet 
and the task of slicing the lump xml results in:
OUTPUT_CII-MWP0002.contigs_singletons.blsatx.viral_Queries.xml
needed to be separated.

This is the module for adding the best hit to the Excel spreadsheet.
'''
""" To work with Excel .xlsx files I am using openpyxl """
from openpyxl import Workbook
from openpyxl import load_workbook
