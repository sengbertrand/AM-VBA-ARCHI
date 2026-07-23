Attribute VB_Name = "DemoRunner"
Option Explicit

Public Sub RunDemo()

    LogInfo "------- DEMO STARTED -------"

    '------------------------------------------------------
    ' 1. Create the initial portfolio
    '------------------------------------------------------
    Dim ptf As Portfolio
    Set ptf = New Portfolio

    ptf.PortfolioId = "EURO_BONDS"

    '------------------------------------------------------
    ' 2. Create the initial OAT position
    '------------------------------------------------------
    Dim posOAT As Position
    Set posOAT = New Position

    posOAT.PortfolioId = ptf.PortfolioId
    posOAT.InstrumentId = "OAT_FR_10Y"
    posOAT.currencyCode = "EUR"
    posOAT.MarketPrice = 99.5
    posOAT.Quantity = 1000000

    ptf.AddPosition posOAT

    '------------------------------------------------------
    ' 3. Create the initial BUND position
    '------------------------------------------------------
    Dim posBUND As Position
    Set posBUND = New Position

    posBUND.PortfolioId = ptf.PortfolioId
    posBUND.InstrumentId = "BUND_DE_10Y"
    posBUND.currencyCode = "EUR"
    posBUND.MarketPrice = 100
    posBUND.Quantity = 500000

    ptf.AddPosition posBUND

    LogInfo "Initial portfolio created : " & ptf.PortfolioId
    LogInfo "Initial position count    : " & ptf.PositionCount

    '------------------------------------------------------
    ' 4. Create an investment decision
    '------------------------------------------------------
    Dim investDec As InvestmentDecision
    Set investDec = New InvestmentDecision

    investDec.DecisionId = "DEC_DEMO_001"
    investDec.PortfolioId = ptf.PortfolioId
    investDec.InstrumentId = "BUND_DE_10Y"
    investDec.TargetWeight = 0.4
    investDec.Rationale = _
        "Increase exposure to German government bonds"
    investDec.DecisionDate = Now

    LogInfo "Investment decision created : " & investDec.DecisionId

    '------------------------------------------------------
    ' 5. Generate the market order
    '------------------------------------------------------
    Dim orderSvc As OrderGenerationService
    Set orderSvc = New OrderGenerationService

    Dim ord As Order
    Set ord = orderSvc.GenerateOrder(investDec)

    LogInfo "Order generated : " & ord.OrderId & _
            " | Instrument = " & ord.InstrumentId & _
            " | Side = " & ord.Side & _
            " | Quantity = " & ord.Quantity

    '------------------------------------------------------
    ' 6. Perform the pre-trade risk assessment
    '------------------------------------------------------
    Dim preTradeRiskSvc As PreTradeRiskService
    Set preTradeRiskSvc = New PreTradeRiskService

    Dim riskReport As PreTradeRiskReport
    Set riskReport = preTradeRiskSvc.EvaluateOrder(ptf, ord)

    Dim riskPresenter As RiskReportPresenter
    Set riskPresenter = New RiskReportPresenter

    riskPresenter.Display riskReport

    LogInfo "Pre-trade risk decision : " & riskReport.Decision
    LogInfo "Current weight = " & _
            Format(riskReport.currentWeight, "0.00%") & _
            " | Simulated weight = " & _
            Format(riskReport.simulatedWeight, "0.00%") & _
            " | Limit weight = " & _
            Format(riskReport.maximumAllowedWeight, "0.00%")

    '------------------------------------------------------
    ' 7. Stop the workflow if the order is rejected
    '------------------------------------------------------
    If Not riskReport.Approved Then

        ord.Status = ORDER_REJECTED

        LogWarn "Order rejected before execution : " & _
                ord.OrderId & _
                " | Reason = " & riskReport.Reason

        Debug.Print "=================="
        Debug.Print "DEMO RESULT"
        Debug.Print "=================="
        Debug.Print "Order Id           : " & ord.OrderId
        Debug.Print "Risk Decision      : " & riskReport.Decision
        Debug.Print "Reason             : " & riskReport.Reason
        Debug.Print "Order Status       : " & ord.Status
        Debug.Print "Portfolio unchanged."

        LogInfo "------- DEMO COMPLETED -------"

        Exit Sub

    End If

    '------------------------------------------------------
    ' 8. Approve and execute the real order
    '------------------------------------------------------
    ord.Status = ORDER_APPROVED

    Dim execSvc As ExecutionService
    Set execSvc = New ExecutionService

    Dim execRep As ExecutionReport
    Set execRep = execSvc.ExecuteOrder(ord)
    
        LogInfo "Order executed : " & ord.OrderId & _
            " | Executed quantity = " & execRep.executedQuantity & _
            " | Execution price = " & execRep.ExecutionPrice

    '------------------------------------------------------
    ' 9. Retrieve the instrument definition
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim currentInstr As instrument
    Set currentInstr = instrumentRepo.GetInstrument(ord.InstrumentId)

    '------------------------------------------------------
    ' 10. Calculate the execution impact
    '------------------------------------------------------
    Dim impactCalc As ExecutionImpactCalculator
    Set impactCalc = New ExecutionImpactCalculator

    Dim execImpact As ExecutionImpact
    Set execImpact = impactCalc.Calculate( _
                        ord:=ord, _
                        report:=execRep, _
                        currentInstr:=currentInstr)

    '------------------------------------------------------
    ' 11. Apply the execution impact to the portfolio
    '------------------------------------------------------
    Dim updateSvc As PortfolioUpdateService
    Set updateSvc = New PortfolioUpdateService

    updateSvc.ApplyExecutionImpact _
        ptf:=ptf, _
        execImpact:=execImpact

    '------------------------------------------------------
    ' 12. Calculate the final portfolio market value
    '------------------------------------------------------
    Dim ptfValSvc As PortfolioValuationService
    Set ptfValSvc = New PortfolioValuationService

    Dim ptfMarketValue As Double
    ptfMarketValue = _
        ptfValSvc.CalculatePortfolioMarketValue(ptf)

    '------------------------------------------------------
    ' 13. Retrieve the final position state
    '------------------------------------------------------
    Dim finalPos As Position
    Set finalPos = ptf.FindPosition(ord.InstrumentId)

    '------------------------------------------------------
    ' 14. Display the final demonstration result
    '------------------------------------------------------
    Debug.Print "=================="
    Debug.Print "DEMO RESULT"
    Debug.Print "=================="
    Debug.Print "Decision Id             : " & investDec.DecisionId
    Debug.Print "Order Id                : " & ord.OrderId
    Debug.Print "Risk Decision           : " & riskReport.Decision
    Debug.Print "Order Status            : " & ord.Status
    Debug.Print "Execution Status        : " & execRep.Status
    Debug.Print "Instrument              : " & ord.InstrumentId
    Debug.Print "Final Quantity          : " & finalPos.Quantity
    Debug.Print "Final Market Price      : " & finalPos.MarketPrice
    Debug.Print "Portfolio Market Value  : " & ptfMarketValue
    Debug.Print "=================="

    LogInfo "Final portfolio market value : " & _
            ptfMarketValue

    LogInfo "------- DEMO COMPLETED -------"

End Sub

