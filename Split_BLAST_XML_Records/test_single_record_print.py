""" OUTDATED SEPARATED INTO TWO DIFFERENT MODULES """
""" I wrote this script for Python 3.X """
""" Author: Lathian J.D. Horton (Jonathan S. Kwiat) Date: 7-24-16

I want to see what the convenience methods 
 
"""




""" Bio.SearchIO is still experimental at this point will throw warning flag,
    which only means the authors consider it changeable but the virtual
    evironment will protect against any changes assuming nobody upgrades
    the bioPython in virtualPython3.5
"""
from Bio import SearchIO



q_g = SearchIO.parse('OUTPUT_Blastn_CII-MWP0002.contigs_singletons.blastx.viral_Queries_against_L_boulardi.xml', 'blast-xml')
count = 0
for query_result in q_g:
    count += 1
    if count == 60:
        break
print(query_result)
print(query_result.hits[0])
print(query_result.hits[0].hsps[0])


    





