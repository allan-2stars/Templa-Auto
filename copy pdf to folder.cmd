@echo off
setlocal
:PROMPT
SET /P AREYOUSURE=Are you sure (Y/[N])?
IF /I "%AREYOUSURE%" NEQ "Y" GOTO END

copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Affinity\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\BASF\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\BCA\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\DAWR\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Dixon\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Fitness First\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Forbes\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Goodlife\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Goodstart\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Jemena\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Jemena Zinfra\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Parkes Shire Council\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\PMC\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\Produce Markets\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\SADHS\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"
copy "C:\Profiles\awang\My Documents\Report Monthly KPI\TK MAXX\Nov-2019\*.pdf" "C:\Profiles\awang\My Documents\Report Monthly KPI\EmailFiles"

:END
endlocal
