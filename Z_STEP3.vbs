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
' Clear the clipboard before starting Excel
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c echo off | clip", 0, True
Dim objExcel, objWorkbook, file, objClipboard, clipboardText, filePath
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
'---------------------------------------------------------------------------------------------
' Attempt to run the macro
Err.Clear
objExcel.Application.Run "Haystack"
If Err.Number <> 0 Then
    LogError("The macro 'Haystack' could not be run. Please check if the macro exists.")
    objWorkbook.Close False  ' Close the workbook without saving
    objExcel.Quit
    WScript.Quit  ' Exit the script
End If
'---------------------------------------------------------------------------------------------
' Save the workbook before closing
objWorkbook.Save
' Close the workbook after saving
objWorkbook.Close False
' Quit Excel
objExcel.Quit

' Run the Python script
Dim pythonPath
pythonPath = "C:\Users\aeros\DIBBSBQ\STEPS\Z_STEP5.py"
If fso.FileExists(pythonPath) Then
    WshShell.Run "python """ & pythonPath & """", 1, True
Else
    LogError("Python script not found at path: " & pythonPath)
End If

Set WshShell = Nothing
Set fso = Nothing