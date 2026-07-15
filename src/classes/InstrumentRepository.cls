VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstrumentRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Repository : InstrumentRepository
'
' Classification:
' Repository
'
' Domain:
' Investment / Security Master
'
' Responsibility:
' Provides reference data for financial instruments.
'
' Data source:
' Instruments worksheet.
'
' Consumer:
' PositionFactory
' ValuationEngine
' Reporting
'
' Mental Model:
' This repository answers:
' "What do we know about this instrument?"
'==========================================================

Public Function GetCurrencyCode(ByVal instrumentId As String) As String

    Dim foundInstrument As Instrument
    Set foundInstrument = GetInstrument(instrumentId)

    GetCurrencyCode = foundInstrument.CurrencyCode

End Function

Public Function GetInstrument(ByVal instrumentId As String) As Instrument

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_INSTRUMENTS)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    Dim row As Long
    Dim foundInstrument As Instrument

    For row = 2 To lastRow

        If Trim(CStr(ws.Cells(row, 1).value)) = instrumentId Then

            Set foundInstrument = New Instrument

            foundInstrument.instrumentId = _
                Trim(CStr(ws.Cells(row, 1).value))

            foundInstrument.InstrumentName = _
                Trim(CStr(ws.Cells(row, 2).value))

            foundInstrument.AssetClass = _
                Trim(CStr(ws.Cells(row, 3).value))

            foundInstrument.CurrencyCode = _
                Trim(CStr(ws.Cells(row, 4).value))

            Set GetInstrument = foundInstrument
            Exit Function

        End If

    Next row

    Err.Raise vbObjectError + 5002, _
              "InstrumentRepository.GetInstrument", _
              "Unknown instrument: " & instrumentId

End Function

