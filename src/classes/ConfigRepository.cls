VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfigRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Repository : ConfigRepository
'
' Classification:
' Repository
'
' Domain:
' Application / Configuration
'
' Responsibility:
' Provides application configuration values stored in
' the Config worksheet.
'
' Data source:
' Config worksheet.
'
' Output:
' Configuration values
'
' Output consumers:
' PreTradeRiskService
' Portfolio services
' Application workflows
'
' Mental Model:
' This repository answers:
' "What configuration value is currently defined for
' this parameter?"
'==========================================================

Private Function GetRawValue( _
                    ByVal parameterName As String) As Variant

    If Trim(parameterName) = vbNullString Then
        Err.Raise vbObjectError + 14001, _
                  "ConfigRepository.GetRawValue", _
                  "Parameter name cannot be empty."
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim row As Long

    For row = 2 To lastRow

        If Trim(CStr(ws.Cells(row, 1).value)) = parameterName Then
            GetRawValue = ws.Cells(row, 2).value
            Exit Function
        End If

    Next row

    Err.Raise vbObjectError + 14002, _
              "ConfigRepository.GetRawValue", _
              "Unknown configuration parameter: " & parameterName

End Function

Public Function GetDoubleValue( _
                    ByVal parameterName As String) As Double

    Dim rawValue As Variant
    rawValue = GetRawValue(parameterName)

    If Not IsNumeric(rawValue) Then
        Err.Raise vbObjectError + 14003, _
                  "ConfigRepository.GetDoubleValue", _
                  "Configuration value is not numeric: " & _
                  parameterName
    End If

    GetDoubleValue = CDbl(rawValue)

End Function

Public Function GetStringValue( _
                    ByVal parameterName As String) As String

    Dim rawValue As Variant
    rawValue = GetRawValue(parameterName)

    GetStringValue = Trim(CStr(rawValue))

End Function

