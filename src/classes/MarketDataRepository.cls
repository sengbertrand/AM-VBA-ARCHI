VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MarketDataRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Repository : MarketDataRepository
'
' Classification:
' Repository
'
' Domain:
' Investment / Market Data
'
' Responsibility:
' Provides current market data for financial instruments.
'
' Data source:
' MarketData worksheet.
'
' Output:
' Market data values
'
' Output consumers:
' PreTradeRiskService
' ValuationEngine
' Reporting
'
' Mental Model:
' This repository answers:
' "What is the current market data for this instrument?"
'==========================================================

Public Function GetMarketPrice(ByVal instrumentId As String) As Double

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MARKET_DATA)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim row As Long
    
    For row = 2 To lastRow
    
        If Trim(CStr(ws.Cells(row, 1).value)) = instrumentId Then
            
            If Not IsNumeric(ws.Cells(row, 2).value) Then
                Err.Raise vbObjectError + 9001, _
                    "MarketDataRepository.GetMarketPrice", _
                    "Invalid market price for instrument : " & instrumentId
            End If
        
            GetMarketPrice = CDbl(ws.Cells(row, 2).value)
            Exit Function
        
        End If
    
    Next row
    
    Err.Raise vbObjectError + 9002, _
        "MarketDataRepository.GetMarketPrice", _
        "%arket price not found for instrument: " & instrumentId
    
End Function
