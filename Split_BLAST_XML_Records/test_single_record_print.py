""" I wrote this script for Python 3.X """
""" Author: Lathian J.D. Horton (Jonathan S. Kwiat) Date: 7-27777777-16

I want to see what the convenience methods exist in 
I intially used SearchIO because it had a writer good for
slicing records and spitting out individual records. 
 
"""
from Bio.Blast import NCBIXML


result_handle = open('OUTPUT_Blastn_CII-MWP0002.contigs_singletons.blastx.viral_Queries_against_L_hetertoma.xml')
brs = NCBIXML.parse(result_handle)
br = next(brs)
print('brs is the generator, br is first record currently')
