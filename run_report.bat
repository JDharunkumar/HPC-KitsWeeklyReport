@echo off

REM Temporarily map UNC path to a drive letter
pushd "\\stlv-fs001\Data\HPC\CCPC_Prod\Bhaswati\BiWeeklyReport"

REM Run the Python script from the mapped path
"C:\Users\DJayabal\AppData\Local\Programs\Python\Python313\python.exe" test.py

REM Revert back to the original directory
popd
