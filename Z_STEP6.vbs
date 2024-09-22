On Error Resume Next ' Enable error handling

' Function to log errors
Sub LogError(errorMessage)
    Dim logFile, logText
    Set logFile = fso.OpenTextFile("C:\Users\aeros\DIBBSBQ\error_log.txt", 8, True)
    logText = Now & " - " & errorMessage & " (Error Code: " & Err.Number & ")"
    logFile.WriteLine logText
    logFile.Close
End Sub

' Create a FileSystemObject for logging
Set fso = CreateObject("Scripting.FileSystemObject")

' Create an instance of Excel
Set objExcel = CreateObject("Excel.Application")

' Check if Excel was created successfully
If Err.Number <> 0 Then
    LogError("Excel could not be started. Please make sure it is installed.")
    WScript.Quit
End If

' Make Excel visible
objExcel.Visible = True

' Disable alerts
objExcel.DisplayAlerts = False

' Attempt to open the workbook
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\aeros\DIBBSBQ\STEPS\BQXL.xlsm")

' Check if the workbook was opened successfully
If Err.Number <> 0 Then
    LogError("The workbook could not be opened. Please check the path and file name.")
    objExcel.Quit
    WScript.Quit
End If

' Attempt to run the macro
Err.Clear
objExcel.Application.Run "PullPrices"
If Err.Number <> 0 Then
    LogError("The macro 'PullPrices' could not be run. Please check if the macro exists.")
    objWorkbook.Close False  ' Close the workbook without saving
    objExcel.Quit
    WScript.Quit  ' Exit the script
End If

' Save the workbook
objWorkbook.Save

