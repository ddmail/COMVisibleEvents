# COMVisibleEvents
Exposing .NET events to COM

# How to debug in Excel 2016
- In Project properties -> Debug -> Start action -> select "Start external program" and add "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
- In Project properties -> Build -> Output -> check "Register for COM interop"
- Compile using "Debug" configuration for "x86" platform
- Hit F5 and wehen Excel starts open "\repos\COMVisibleEvents\VBA-Client\Excel-VBA-Client.xlsm"
- Set breakpoints and in Excel sheet click the button...
