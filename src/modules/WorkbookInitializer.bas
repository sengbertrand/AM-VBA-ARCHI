Attribute VB_Name = "WorkbookInitializer"
Option Explicit

'=========================================================
' Module : WorkbookInitializer
' Purpose :
'   Creates and initializes all worksheets required by the
'   Asset Management demo application.
'=========================================================

Public Sub InitializeWorkbook()

    EnsureSheet SHEET_LOGS
    EnsureSheet SHEET_CONFIG
    EnsureSheet SHEET_POSITIONS
    EnsureSheet SHEET_TRADES
    EnsureSheet SHEET_INSTRUMENTS
    EnsureSheet SHEET_MARKET_DATA
    EnsureSheet SHEET_RISK_REPORT
    
    InitializeLogsSheet
    InitializeConfigSheet
    InitializePositionsSheet
    InitializeTradesSheet
    InitializeInstrumentsSheet
    InitializeMarketDataSheet
    InitializeRiskReportSheet
    
End Sub

Private Sub EnsureSheet(ByVal sheetName As String)

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
End Sub

Private Sub InitializeLogsSheet()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LOGS)
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "Timestamp"
    ws.Cells(1, 2).value = "Level"
    ws.Cells(1, 3).value = "Message"
    
End Sub

Private Sub InitializeConfigSheet()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "Parameter"
    ws.Cells(1, 2).value = "Value"
    
    ws.Cells(2, 1).value = "BaseCurrency"
    ws.Cells(2, 2).value = "EUR"
    
End Sub

Private Sub InitializePositionsSheet()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_POSITIONS)
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "InstrumentId"
    ws.Cells(1, 2).value = "Quantity"
    ws.Cells(1, 3).value = "Currency"
    ws.Cells(1, 4).value = "Price"
    
End Sub

Private Sub InitializeTradesSheet()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_TRADES)
    
    ws.Cells.Clear
    ws.Cells(1, 1).value = "TradeId"
    ws.Cells(1, 2).value = "OrderId"
    ws.Cells(1, 3).value = "PortfolioId"
    ws.Cells(1, 4).value = "InstrumentId"
    ws.Cells(1, 5).value = "Side"
    ws.Cells(1, 6).value = "Quantity"
    ws.Cells(1, 7).value = "Price"
    ws.Cells(1, 8).value = "Currency"
    
End Sub

Private Sub InitializeInstrumentsSheet()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_INSTRUMENTS)

    ws.Cells.Clear

    ws.Cells(1, 1).value = "InstrumentId"
    ws.Cells(1, 2).value = "InstrumentName"
    ws.Cells(1, 3).value = "AssetClass"
    ws.Cells(1, 4).value = "CurrencyCode"

End Sub

Private Sub InitializeMarketDataSheet()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MARKET_DATA)
    
    ws.Cells.Clear
    
    ws.Cells(1, 1).value = "InstrumentId"
    ws.Cells(1, 2).value = "MarketPrice"
    ws.Cells(1, 3).value = "PriceTimeStamp"
    
End Sub

Private Sub InitializeRiskReportSheet()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_RISK_REPORT)
    
    ws.Cells.Clear
    
    ws.Cells(1, 1).value = "PRE-TRADE RISK REPORT"
    
End Sub
