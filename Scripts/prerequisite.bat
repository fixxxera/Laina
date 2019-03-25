@echo off
echo Downloading Python
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://www.python.org/ftp/python/3.6.3/python-3.6.3-amd64.exe', ' C:\Users\%username%\Documents\python3.6.2.exe') }"
echo Installing Python
C:\Users\%username%\Documents\python3.6.2.exe /quiet InstallAllUsers=1 TargetDir='%ProgramFiles%\Python36' Include_pip=1 Include_test=0 PrependPath=1
echo Installing modules
"%ProgramFiles%\Python36\Scripts\pip.exe" install xlrd xlsxwriter lxml bs4 requests selenium pyqt5
SET firepath="http://download-origin.cdn.mozilla.net/pub/firefox/releases/56.0/win64/en-US/Firefox Setup 56.0.exe"
echo Downloading Firefox 56
powershell -command "& { (New-Object Net.WebClient).DownloadFile('%firepath%', 'C:\Users\%username%\Documents\Firefox.exe') }"
echo Installing Firefox 56
C:\Users\%username%\Documents\Firefox.exe /S
echo Downloading Geckodriver
powershell -command "& { (New-Object Net.WebClient).DownloadFile('https://github.com/mozilla/geckodriver/releases/download/v0.19.0/geckodriver-v0.19.0-win64.zip', 'C:\Users\%username%\Documents\geckodriver.zip') }"
echo Downloading 7zip
powershell -command "& { (New-Object Net.WebClient).DownloadFile('http://www.7-zip.org/a/7z1701-x64.exe', 'C:\Users\%username%\Documents\7z.exe') }"
echo Installing 7zip
"C:\Users\%username%\Documents\7z.exe" /S /D="%ProgramFiles%\7-Zip"
echo unzipping Geckodriver
"%ProgramFiles%\7-Zip\7z.exe" e "C:\Users\%username%\Documents\geckodriver.zip" -o"%ProgramFiles%\Python36\" > NUL:
echo Creating folder "C:\Users\%username%\Dropbox\XLSX"
mkdir "C:\Users\%username%\Dropbox\XLSX"
Echo Done
pause


