@ECHO OFF
copy data.xlsx backup_data.xlsx /y
copy Immunizations.xlsx backup_immunizations.xlsx /y
git reset --hard
git pull
PAUSE