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

Private Sub WriteLog(ByVal level As String, ByVal message As String)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOGS)
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    
    ws.Cells(nextRow, 1).value = Now
    ws.Cells(nextRow, 2).value = level
    ws.Cells(nextRow, 3).value = message
    
End Sub
