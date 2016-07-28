@echo off

rem THIS LINE OPENS MULTIPLE INTERNET EXPLORER LINKS (IN TABS), USING A SEPARATE VB SCRIPT
cscript //nologo ie_tabs2.vbs

rem OPEN INTERACTION DESKTOP
start "" "C:\Program Files (x86)\Interactive Intelligence\ICUserApps\InteractionDesktop.exe"


rem OPEN TELETRACKING TRANSFERCENTER MODULE
start "" "C:\Program Files (x86)\Teletracking\TransferCenter\TransferCenter.exe"

rem OPEN ONENOTE
start "C:\Program Files (x86)\Microsoft Office\Office14\ONENOTE.EXE" "G:\Patient_Throughput\~Departments\Transfer Center and Bed Board\Open Notebook.onetoc2" 
"G:\Patient_Throughput\~Departments\Transfer Center and Bed Board\Open Notebook.onetoc2"

rem START ACTIVE DASHBOARD
start "" "C:\Program Files\Healthware Systems\ActiveDASHBOARD\Dashboard.exe"

rem START OUTLOOK
start "" "C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"

rem OPEN INQUICKER IN A GOOGLE CHROME WINDOW
start chrome https://console.inquicker.com/ascensionhealth
