# sportfprall
pyinstaller --noconfirm --onefile --windowed SportForAll.py

з папкою
pyinstaller --noconfirm --onefile --windowed --add-data "templates;templates" SportForAll.py

pyinstaller --noconfirm --onefile --windowed ^
--add-data "templates;templates" ^
--add-data "ШАБЛОН_кошторис_розумний.xlsx;." ^
SportForAll.py
