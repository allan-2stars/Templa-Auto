@echo off
setlocal
:PROMPT
SET /P AREYOUSURE=Are you sure (Y/[N])?
IF /I "%AREYOUSURE%" NEQ "Y" GOTO END

move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Affinity*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Affinity\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*BASF*" "C:\Profiles\awang\My Documents\Report Monthly KPI\BASF\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Bowling*" "C:\Profiles\awang\My Documents\Report Monthly KPI\BCA\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*DAWR*" "C:\Profiles\awang\My Documents\Report Monthly KPI\DAWR\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Dixon*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Dixon\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Fitness*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Fitness First\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Forbes*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Forbes\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Goodlife*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Goodlife\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Goodstart*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Goodstart\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Jemena KPI*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Jemena\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Zinfra*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Jemena Zinfra\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Parkes*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Parkes Shire Council\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*PMC*" "C:\Profiles\awang\My Documents\Report Monthly KPI\PMC\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*Produce*" "C:\Profiles\awang\My Documents\Report Monthly KPI\Produce Markets\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*DHS*" "C:\Profiles\awang\My Documents\Report Monthly KPI\SADHS\Dec-2019"
move "C:\Profiles\awang\My Documents\Report Monthly KPI\MasterFiles\*MAXX*" "C:\Profiles\awang\My Documents\Report Monthly KPI\TK MAXX\Dec-2019"


:END
endlocal