Attribute VB_Name = "TestRisk"
Option Explicit

Public Sub TestRiskEngine()
    
    '------------------------------
    ' Create portfolio
    '------------------------------
    
    Dim ptf As Portfolio
    Set ptf = New Portfolio

    ptf.PortfolioId = "EURO_BONDS"
    
    '--------------------------------
    ' Create OAT position
    '-------------------------------
    
    Dim posOAT As Position
    Set posOAT = New Position
    
    posOAT.PortfolioId = "EURO_BONDS"
    posOAT.instrumentId = "OAT_FR_10Y"
    posOAT.CurrencyCode = "EUR"
    posOAT.MarketPrice = 99.5
    posOAT.Quantity = 1000000
    
    ptf.AddPosition posOAT
    
    '------------------------------
    ' Create BUND position
    '----------------------
    
    Dim posBUND As Position
    Set posBUND = New Position
    
    posBUND.PortfolioId = "EURO_BONDS"
    posBUND.instrumentId = "BUND_DE_10Y"
    posBUND.CurrencyCode = "EUR"
    posBUND.MarketPrice = 100
    posBUND.Quantity = 500000
    
    ptf.AddPosition posBUND
    
    '-------------------------
    ' Calculate position weights
    '-------------------------
    
    Dim riskSvc As RiskEngine
    Set riskSvc = New RiskEngine
    
    
    Dim oatWeight As Double
    oatWeight = riskSvc.CalculatePositionWeight(ptf, posOAT)
    
    Dim bundWeight As Double
    bundWeight = riskSvc.CalculatePositionWeight(ptf, posBUND)
    
    '--------------------------
    ' Verify results
    '------------------
    
    Debug.Print "Portfolio Id : " & ptf.PortfolioId
    Debug.Print "OAT Weight : " & Format(oatWeight, "0.00%")
    Debug.Print "BUND Weight : " & Format(bundWeight, "0.00%")
    Debug.Print "Total Weight : " & Format(oatWeight + bundWeight, "0.00%")
    
End Sub

Public Sub TestPreTradeRiskService()
    
    '-------------------------
    ' Create the real portfolio
    '-----------------------
    
    Dim ptf As Portfolio
    Set ptf = New Portfolio
    
    ptf.PortfolioId = "EURO_BONDS"
    
    '-------------------------
    ' Create OAT position
    '---------------------------
    
    Dim posOAT As Position
    Set posOAT = New Position
    
    posOAT.PortfolioId = ptf.PortfolioId
    posOAT.instrumentId = "OAT_FR_10Y"
    posOAT.CurrencyCode = "EUR"
    posOAT.Quantity = 1000000
    posOAT.MarketPrice = 99.5
    
    ptf.AddPosition posOAT
    
    '--------------------------
    ' Create BUND position
    '---------------------
    
    Dim posBUND As Position
    Set posBUND = New Position
    
    posBUND.PortfolioId = ptf.PortfolioId
    posBUND.instrumentId = "BUND_DE_10Y"
    posBUND.CurrencyCode = "EUR"
    posBUND.MarketPrice = 100
    posBUND.Quantity = 500000
    
    ptf.AddPosition posBUND
    
    '--------------------
    ' Create pre-trade risk service
    '-----------------------
    
    Dim preTradeRiskSvc As PreTradeRiskService
    Set preTradeRiskSvc = New PreTradeRiskService
    
    '==================
    ' Scenario 1: expected approval
    '======================
    
    Dim approvedOrd As Order
    Set approvedOrd = New Order
    
    approvedOrd.OrderId = "ORD_RISK_APPROVED"
    approvedOrd.PortfolioId = ptf.PortfolioId
    approvedOrd.instrumentId = "BUND_DE_10Y"
    approvedOrd.Side = SIDE_BUY
    approvedOrd.Quantity = 100000
    approvedOrd.Status = ORDER_CREATED
    
    Dim approvedReport As PreTradeRiskReport
    Set approvedReport = preTradeRiskSvc.EvaluateOrder(ptf, approvedOrd)
    
    Debug.Print "---------------------------"
    Debug.Print "APPROVAL_SCENARIO"
    Debug.Print "---------------------------"
    Debug.Print "Decision           : " & approvedReport.Decision
    Debug.Print "Approved         : " & approvedReport.Approved
    Debug.Print "Current weight : " & _
                                Format(approvedReport.CurrentWeight, "0.00%")
    Debug.Print "Simulated weight : " & _
                                Format(approvedReport.SimulatedWeight, "0.00%")
    Debug.Print "Maximum Weight : " & _
                                Format(approvedReport.MaximumAllowedWeight, "0.00%")
    Debug.Print "Reason                 : " & approvedReport.Reason
    Debug.Print "Snapshot Id          : " & approvedReport.SnapshotId
    
    '============================
    ' Scenario 2: expected rejection
    '====================
    
    Dim rejectedOrd As Order
    Set rejectedOrd = New Order
    
    rejectedOrd.OrderId = "ORDER_RISK_REJECTED"
    rejectedOrd.PortfolioId = ptf.PortfolioId
    rejectedOrd.instrumentId = "BUND_DE_10Y"
    rejectedOrd.Side = SIDE_BUY
    rejectedOrd.Quantity = 3000000
    rejectedOrd.Status = ORDER_CREATED
    
    Dim rejectedReport As PreTradeRiskReport
    Set rejectedReport = preTradeRiskSvc.EvaluateOrder(ptf, rejectedOrd)
    
    Debug.Print "--------------------"
    Debug.Print "REJECTION SCENARIO"
    Debug.Print "---------------------"
    Debug.Print "Decision               : " & rejectedReport.Decision
    Debug.Print "Approved             : " & rejectedReport.Approved
    Debug.Print "Current weight     : " & _
                                        Format(rejectedReport.CurrentWeight, "0.00%")
    Debug.Print "Simulated weight : " & _
                                        Format(rejectedReport.SimulatedWeight, "0.00%")
    Debug.Print "Maximum Allowed weight : " & _
                            Format(rejectedReport.MaximumAllowedWeight, "0.00%")
    Debug.Print "Reason                  : " & rejectedReport.Reason
    Debug.Print "Snapshort              : " & rejectedReport.SnapshotId
    
    '---------------------------------------
    ' Verify that the real portfolio was not modified
    '---------------------------------------
    
    Debug.Print "-------------"
    Debug.Print "REAL PORTFOLIO CONTROL"
    Debug.Print "---------------"
    Debug.Print "Real BUND Qty          : "; posBUND.Quantity
    
    
End Sub
