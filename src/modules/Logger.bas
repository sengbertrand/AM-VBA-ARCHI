Attribute VB_Name = "Logger"
Option Explicit

'=========================================================
' Module : Logging
' Purpose :
'   Centralizes application logging.
'
' Responsibilities:
'   - Record application events.
'   - Record warnings and errors.
'   - Persist log entries into the Logs worksheet.
'
' Dependencies:
'   - Constants
'   - Logs worksheet
'=========================================================


Public Sub LogInfo(ByVal message As String)
    WriteLog LOG_INFO, message
End Sub

Public Sub LogWarn(ByVal message As String)
    WriteLog LOG_WARN, message
End Sub

Public Sub LogError(ByVal message As String)
    WriteLog LOG_ERROR, message
End Sub


Private Sub WriteLog1( _
    ByVal level As String, _
    ByVal message As String)

    Debug.Print "WriteLog - Start"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOGS)

    Debug.Print "WriteLog - Worksheet found"

    Dim lastUsedRow As Long
    lastUsedRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Debug.Print "WriteLog - Last row = " & lastUsedRow

    Dim nextRow As Long
    nextRow = lastUsedRow + 1

    Debug.Print "WriteLog - Next row = " & nextRow

    ws.Cells(nextRow, 1).value = Now
    Debug.Print "WriteLog - Date written"

    ws.Cells(nextRow, 2).value = level
    Debug.Print "WriteLog - Level written"

    ws.Cells(nextRow, 3).value = message
    Debug.Print "WriteLog - Message written"

End Sub
Private Sub WriteLog(ByVal level As String, ByVal message As String)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOGS)
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = level
    ws.Cells(nextRow, 3).value = message
    
End Sub
