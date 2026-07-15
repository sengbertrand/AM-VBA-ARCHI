Attribute VB_Name = "DemoRunner"
Option Explicit

Public Sub RunDemo()
    
    LogInfo "-------DEMO STARTED -------------"
    
    '-------------------
    ' 1. Create the initial portfolio
    '-------------
    
    Dim ptf As Portfolio
    Set ptf = New Portfolio
    
    ptf.PortfolioId = "EURO_BONDS"
    
    
    '-------------
    ' 2. Create the initial OAT position
    '-----------
    
    Dim posOAT As Position
    Set posOAT = New Position
    
    posOAT.PortfolioId = ptf.PortfolioId
    posOAT.instrumentId = "OAT_FR_10Y"
    posOAT.CurrencyCode = "EUR"
    posOAT.MarketPrice = 99.5
    posOAT.Quantity = 1000000
    
    ptf.AddPosition posOAT
    
    '--------------------------
    '3. Create the initial BUND position
    '--------------
    
    Dim posBUND As Position
    Set posBUND = New Position
    
    posBUND.PortfolioId = ptf.PortfolioId
    posBUND.instrumentId = "BUND_DE_10Y"
    posBUND.CurrencyCode = "EUR"
    posBUND.MarketPrice = 100
    posBUND.Quantity = 500000
    
    ptf.AddPosition posBUND

    LogInfo "Initial portfolio created : " & ptf.PortfolioId
    LogInfo "Initial position count     : " & ptf.PositionCount
    
    '----------------------
    ' 4. Create an investment decision
    '-----------------
    
    Dim investDec As InvestmentDecision
    Set investDec = New InvestmentDecision
    
    investDec.DecisionId = "DEC_DEMO_001"
    investDec.PortfolioId = ptf.PortfolioId
    investDec.instrumentId = "BUND_DE_10Y"
    investDec.TargetWeight = 0.4
    investDec.Rationale = _
    "Increase exposure to German government bonds"
    investDec.DecisionDate = Now
    
    
    LogInfo "Investment decision created : " & investDec.DecisionId
    
    '-----------
    ' 5. Generate the market order
    '-------------------
    
    Dim orderSvc As OrderGenerationService
    Set orderSvc = New OrderGenerationService
    
    Dim ord As Order
    Set ord = orderSvc.GenerateOrder(investDec)
    
    LogInfo "Order generated : " & ord.OrderId & _
        " | Instrument : " & ord.instrumentId & _
        " | Side = " & ord.Side & _
        " | Quantity = " & ord.Quantity
        
    '----------------
    ' 6. Perform the pre-trade risk assessment
    '--------------------
    
    Dim preTradeRiskSvc As PreTradeRiskService
    Set preTradeRiskSvc = New PreTradeRiskService
    
    Dim riskReport As PreTradeRiskReport
    Set riskReport = preTradeRiskSvc.EvaluateOrder(ptf, ord)
    
    Dim riskPresenter As RiskReportPresenter
    Set riskPresenter = New RiskReportPresenter
    
    riskPresenter.Display riskReport
    
    LogInfo "Pre-trade risk decision : " & riskReport.Decision
    LogInfo "Current weight = " & _
        Format(riskReport.CurrentWeight, "0.00%") & _
        " | Simulated weight = " & _
        Format(riskReport.SimulatedWeight, "0.00%") & _
        " | Limit weight = " & _
        Format(riskReport.MaximumAllowedWeight, "0.00%")
    
    '------------
    ' 7. Stop the workflow if order is rejected
    '--------------------
    
    If Not riskReport.Approved Then
        
        ord.Status = ORDER_REJECTED
        
        LogWarn "Order rejected before execution: " & _
            ord.OrderId & _
            " | Reason = " & riskReport.Reason
        
        Debug.Print "============"
        Debug.Print "DEMO RESULT"
        Debug.Print "============"
        Debug.Print "Order Id           : " & ord.OrderId
        Debug.Print "Risk Decision   : " & riskReport.Decision
        Debug.Print "Reason            : " & riskReport.Reason
        Debug.Print "Order Status    : " & ord.Status
        Debug.Print "Portfolio is not modified."
        
        LogInfo "--------- DEMO COMPLETED ------------"
        
        Exit Sub
    
    End If
    
    '-------------------
    '8. Approve and execute the real order
    '---------------------
    
    ord.Status = ORDER_APPROVED
    
    Dim execSvc As ExecutionService
    Set execSvc = New ExecutionService
    
    Dim execRep As ExecutionReport
    Set execRep = execSvc.ExecuteOrder(ord)
    
    LogInfo "Order executed : " & ord.OrderId & _
        " | Executed quantity = " & execRep.ExecutedQuantity & _
        " | execution price = " & execRep.ExecutionPrice
    
    '-----------------------
    ' 9. Apply the real execution to the portfolio
    '------------------
    
    Dim updateSvc As PortfolioUpdateService
    Set updateSvc = New PortfolioUpdateService
    
    updateSvc.ApplyExecution ptf, ord, execRep
    
    '----------------------
    ' 10. Calculate final portfolio market value
    '----------------------------
    
    Dim ptfValSvc As PortfolioValuationService
    Set ptfValSvc = New PortfolioValuationService
    
    Dim ptfMarketValue As Double
    ptfMarketValue = ptfValSvc.CalculatePortfolioMarketValue(ptf)
    
    '-----------
    ' 11. Retrieve the final position state
    '----------------------
    
    Dim finalPos As Position
    Set finalPos = ptf.FindPosition(ord.instrumentId)
    
    '------------------
    ' 12. Display the final demonstration result
    '----------------------
    
    Debug.Print "==========="
    Debug.Print "DEMO RESULT"
    Debug.Print "==========="
    Debug.Print "Decision Id                    : " & investDec.DecisionId
    Debug.Print "Order Id                        : " & ord.OrderId
    Debug.Print "Risk Decision                : " & riskReport.Decision
    Debug.Print "Order Status                  : " & ord.Status
    Debug.Print "Execution Status            : " & execRep.Status
    Debug.Print "Instrument                     : " & ord.instrumentId
    Debug.Print "Final Quantity                 : " & finalPos.Quantity
    Debug.Print "Final MP                         : " & finalPos.MarketPrice
    Debug.Print "Portfolio MV                   : " & ptfMarketValue
    Debug.Print "============"
    
    LogInfo "Final portfolio market value : " & _
            ptfMarketValue
    
    LogInfo "-------- DEMO COMPLET ---------"
End Sub
