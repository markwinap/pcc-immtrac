@ECHO OFF
copy data.xlsx backup/data.xlsx /y
copy Immunizations.xlsx backup/Immunizations.xlsx /y
git reset --hard
git pull
PAUSE