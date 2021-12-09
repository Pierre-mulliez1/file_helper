@echo off
cd C:\Users\pierre.a.mulliez\OneDrive - Accenture\Documents\Python_helpers
rm clean env 
del excels /q /s
rm insert here valid code coding environment
pip install .
echo %CD%
rm making sure the necessary module are downloaded 
pip install -r requirements.txt
python -c "import helper; from helper import to_excel; to_excel()"
rm all dataset to backup folder
MOVE /Y data\* backup
del data /q /s
rem end 
