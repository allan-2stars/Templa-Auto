from functions.functions_kpi_analysis import KPI_Analysis
from functions.functions_kpi_failedQA import KPI_FaildedQA


## run site re-allocation and update qa recipients all together.
KPI_Analysis()
KPI_FaildedQA()