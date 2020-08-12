from functions.functions_site import Site_Reassign
from functions.functions_qa import QA_Recipients


## run site re-allocation and update qa recipients all together.
#
## write all logs to file
import sys
sys.stdout=open("All-CSM-Reallocate.txt","w")

Site_Reassign()
QA_Recipients()

## close file handle
sys.stdout.close()
