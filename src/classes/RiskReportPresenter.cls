VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiskReportPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Display(ByRef riskReport As PreTradeRiskReport)
    
    '-----------
    ' pre-conditions'
    '----------------
    
    If riskReport Is Nothing Then
        Err.Raise vbObjectError + 15001, _
            "RiskReportPresenter.Display", _
            "PreTradeRiskReport cannont be Nothing."
    End If
    
    '------------------
    ' Access the presentation worksheet
    '--------------------
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_RISK_REPORT)
    
    '------------
    ' Clear the previous displayed report
    '-----------------
    
    ws.Cells.Clear
    
    '--------------------
    ' Display report title
    '-------------
    
    ws.Range("A1:B1").Merge
    
    ws.Cells(1, 1).value = "PRE-TRADE RISK  REPORT"
    
    With ws.Range("A1:B1")
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ws.Rows(1).RowHeight = 26

    '---------
    ' Display report identification
    '-----------------
    
    ws.Cells(3, 1).value = "Report Id"
    ws.Cells(3, 2).value = riskReport.ReportId
    
    ws.Cells(4, 1).value = "Portfolio Id"
    ws.Cells(4, 2).value = riskReport.PortfolioId
    
    ws.Cells(5, 1).value = "Order Id"
    ws.Cells(5, 2).value = riskReport.OrderId
    
    ws.Cells(6, 1).value = "Instrument Id"
    ws.Cells(6, 2).value = riskReport.instrumentId
    
    ws.Cells(7, 1).value = "Snapshot Id"
    ws.Cells(7, 2).value = riskReport.SnapshotId
    
    ws.Cells(8, 1).value = "Decision Time"
    ws.Cells(8, 2).value = riskReport.DecisionTime
    
    '--------------
    ' Display risk measures
    '--------------------
    
    ws.Cells(10, 1).value = "Current Weight"
    ws.Cells(10, 2).value = riskReport.CurrentWeight
    
    ws.Cells(11, 1).value = "Simulated Weight"
    ws.Cells(11, 2).value = riskReport.SimulatedWeight
    
    ws.Cells(12, 1).value = "Maximum Allowed Weight"
    ws.Cells(12, 2).value = riskReport.MaximumAllowedWeight
    
    '-----------------
    ' Display final decision
    '-------------------
    
    ws.Cells(14, 1).value = "Decision"
    ws.Cells(14, 2).value = riskReport.Decision
    
    ws.Cells(15, 1).value = "Approved"
    ws.Cells(15, 2).value = riskReport.Approved
    
    ws.Cells(16, 1).value = "Reason"
    ws.Cells(16, 2).value = riskReport.Reason
    
    '------------
    ' Format labels
    '-----------------
    
    With ws.Range("A3:A8")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With
    
    With ws.Range("A10:A12")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With
    
    With ws.Range("A14:A16")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With

    
    '----------------
    ' Apply number formats
    '---------------------
    
    ws.Cells(8, 2).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Cells(10, 2).NumberFormat = "0.00%"
    ws.Cells(11, 2).NumberFormat = "0.00%"
    ws.Cells(12, 2).NumberFormat = "0.00%"
    
    '------------------
    'Highlight simulated weight
    '-----------------------
    
    If riskReport.SimulatedWeight <= _
        riskReport.MaximumAllowedWeight Then
        
        ws.Cells(11, 2).Interior.Color = RGB(198, 239, 206)
        ws.Cells(11, 2).Font.Color = RGB(0, 97, 0)
        
    Else
            
        ws.Cells(11, 2).Interior.Color = RGB(255, 199, 206)
        ws.Cells(11, 2).Font.Color = RGB(156, 0, 6)
    
    End If
    
    '--------------
    ' Highlight Final decision
    '----------------------
    
    With ws.Cells(14, 2)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        
        If riskReport.Approved Then
            .Interior.Color = RGB(0, 176, 80)
        Else
            .Interior.Color = RGB(192, 0, 0)
        End If
    End With
    
    '----------------------
    ' Format borders and alignment
    '---------------------------
    With ws.Range("A3:B16").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ws.Range("A3:A16").HorizontalAlignment = xlLeft
    ws.Range("B3:B16").HorizontalAlignment = xlLeft
    
    ws.Cells(10, 2).HorizontalAlignment = xlRight
    ws.Cells(11, 2).HorizontalAlignment = xlRight
    ws.Cells(12, 2).HorizontalAlignment = xlRight
    

    '-----------
    'improve readability
    '------------------
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 60
    
    ws.Cells(16, 2).WrapText = True
    ws.Rows(16).AutoFit

    
End Sub
