Attribute VB_Name = "TestPortfolioWorkflow"
Option Explicit

Public Sub TestPortfolioUpdateService()

    '------------------------------------------------------
    ' 1. Create the initial portfolio
    '------------------------------------------------------
    Dim ptf As Portfolio
    Set ptf = New Portfolio

    ptf.PortfolioId = "EURO_BONDS"

    '------------------------------------------------------
    ' 2. Create the initial position
    '------------------------------------------------------
    Dim pos As Position
    Set pos = New Position

    pos.PortfolioId = ptf.PortfolioId
    pos.InstrumentId = "OAT_FR_10Y"
    pos.Quantity = 1000000
    pos.MarketPrice = 99.5
    pos.currencyCode = "EUR"

    ptf.AddPosition pos

    Debug.Print "Initial Qty : " & pos.Quantity

    '------------------------------------------------------
    ' 3. Create an executed order
    '------------------------------------------------------
    Dim ord As Order
    Set ord = New Order

    ord.OrderId = "ORD_002"
    ord.PortfolioId = ptf.PortfolioId
    ord.InstrumentId = "OAT_FR_10Y"
    ord.Side = SIDE_BUY
    ord.Quantity = 300000
    ord.Status = ORDER_CREATED

    '------------------------------------------------------
    ' 4. Execute the order
    '------------------------------------------------------
    Dim execSvc As ExecutionService
    Set execSvc = New ExecutionService

    Dim execRep As ExecutionReport
    Set execRep = execSvc.ExecuteOrder(ord)

    '------------------------------------------------------
    ' 5. Retrieve the instrument definition
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim currentInstr As instrument
    Set currentInstr = instrumentRepo.GetInstrument(ord.InstrumentId)

    '------------------------------------------------------
    ' 6. Calculate the execution impact
    '------------------------------------------------------
    Dim impactCalc As ExecutionImpactCalculator
    Set impactCalc = New ExecutionImpactCalculator

    Dim execImpact As ExecutionImpact
    Set execImpact = impactCalc.Calculate( _
                        ord:=ord, _
                        report:=execRep, _
                        currentInstr:=currentInstr)

    '------------------------------------------------------
    ' 7. Apply the execution impact
    '------------------------------------------------------
    Dim updateSvc As PortfolioUpdateService
    Set updateSvc = New PortfolioUpdateService

    updateSvc.ApplyExecutionImpact _
        ptf:=ptf, _
        execImpact:=execImpact

    '------------------------------------------------------
    ' 8. Verify the updated position
    '------------------------------------------------------
    Debug.Print "Final Qty   : " & pos.Quantity
    Debug.Print "Final Price : " & pos.MarketPrice

End Sub

Public Sub TestEndToEndWorkflow()

    '------------------------------------------------------
    ' 1. Create the initial portfolio
    '------------------------------------------------------
    Dim ptf As Portfolio
    Set ptf = New Portfolio

    ptf.PortfolioId = "EURO_BONDS"

    '------------------------------------------------------
    ' 2. Create the initial position
    '------------------------------------------------------
    Dim pos As Position
    Set pos = New Position

    pos.PortfolioId = ptf.PortfolioId
    pos.InstrumentId = "OAT_FR_10Y"
    pos.Quantity = 1000000
    pos.MarketPrice = 99.5
    pos.currencyCode = "EUR"

    ptf.AddPosition pos

    Debug.Print "Initial Qty : " & pos.Quantity

    '------------------------------------------------------
    ' 3. Create the investment decision
    '------------------------------------------------------
    Dim investDec As InvestmentDecision
    Set investDec = New InvestmentDecision

    investDec.DecisionId = "DEC_002"
    investDec.PortfolioId = ptf.PortfolioId
    investDec.InstrumentId = "OAT_FR_10Y"
    investDec.TargetWeight = 0.15
    investDec.DecisionDate = Date
    investDec.Rationale = _
        "Increase exposure to French government bonds"

    '------------------------------------------------------
    ' 4. Generate the order
    '------------------------------------------------------
    Dim orderSvc As OrderGenerationService
    Set orderSvc = New OrderGenerationService

    Dim ord As Order
    Set ord = orderSvc.GenerateOrder(investDec)

    Debug.Print "Generated Order : " & ord.OrderId
    Debug.Print "Order Qty       : " & ord.Quantity
    Debug.Print "Order Status    : " & ord.Status

    '------------------------------------------------------
    ' 5. Execute the order
    '------------------------------------------------------
    Dim execSvc As ExecutionService
    Set execSvc = New ExecutionService

    Dim execRep As ExecutionReport
    Set execRep = execSvc.ExecuteOrder(ord)

    Debug.Print "Execution Status : " & execRep.Status
    Debug.Print "Execution Price  : " & execRep.ExecutionPrice

    '------------------------------------------------------
    ' 6. Retrieve the instrument definition
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim currentInstr As instrument
    Set currentInstr = instrumentRepo.GetInstrument(ord.InstrumentId)

    '------------------------------------------------------
    ' 7. Calculate the execution impact
    '------------------------------------------------------
    Dim impactCalc As ExecutionImpactCalculator
    Set impactCalc = New ExecutionImpactCalculator

    Dim execImpact As ExecutionImpact
    Set execImpact = impactCalc.Calculate( _
                        ord:=ord, _
                        report:=execRep, _
                        currentInstr:=currentInstr)

    '------------------------------------------------------
    ' 8. Apply the execution impact
    '------------------------------------------------------
    Dim updateSvc As PortfolioUpdateService
    Set updateSvc = New PortfolioUpdateService

    updateSvc.ApplyExecutionImpact _
        ptf:=ptf, _
        execImpact:=execImpact

    '------------------------------------------------------
    ' 9. Verify the final portfolio state
    '------------------------------------------------------
    Debug.Print "Final Qty   : " & pos.Quantity
    Debug.Print "Final Price : " & pos.MarketPrice

End Sub

Public Sub TestNewPositionCreation()

    '------------------------------------------------------
    ' 1. Create an empty portfolio
    '------------------------------------------------------
    Dim ptf As Portfolio
    Set ptf = New Portfolio

    ptf.PortfolioId = "EURO_BONDS"

    '------------------------------------------------------
    ' 2. Create an order for a new instrument
    '------------------------------------------------------
    Dim ord As Order
    Set ord = New Order

    ord.OrderId = "ORD_003"
    ord.PortfolioId = ptf.PortfolioId
    ord.InstrumentId = "BUND_DE_10Y"
    ord.Side = SIDE_BUY
    ord.Quantity = 500000
    ord.Status = ORDER_CREATED

    '------------------------------------------------------
    ' 3. Execute the order
    '------------------------------------------------------
    Dim execSvc As ExecutionService
    Set execSvc = New ExecutionService

    Dim execRep As ExecutionReport
    Set execRep = execSvc.ExecuteOrder(ord)

    '------------------------------------------------------
    ' 4. Retrieve the instrument definition
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim currentInstr As instrument
    Set currentInstr = instrumentRepo.GetInstrument(ord.InstrumentId)

    '------------------------------------------------------
    ' 5. Calculate the execution impact
    '------------------------------------------------------
    Dim impactCalc As ExecutionImpactCalculator
    Set impactCalc = New ExecutionImpactCalculator

    Dim execImpact As ExecutionImpact
    Set execImpact = impactCalc.Calculate( _
                        ord:=ord, _
                        report:=execRep, _
                        currentInstr:=currentInstr)

    '------------------------------------------------------
    ' 6. Apply the execution impact
    '------------------------------------------------------
    Dim updateSvc As PortfolioUpdateService
    Set updateSvc = New PortfolioUpdateService

    updateSvc.ApplyExecutionImpact _
        ptf:=ptf, _
        execImpact:=execImpact

    '------------------------------------------------------
    ' 7. Retrieve the newly created position
    '------------------------------------------------------
    Dim pos As Position
    Set pos = ptf.FindPosition("BUND_DE_10Y")

    '------------------------------------------------------
    ' 8. Verify the portfolio state
    '------------------------------------------------------
    Debug.Print "Position Count : " & ptf.PositionCount
    Debug.Print "Instrument     : " & pos.InstrumentId
    Debug.Print "Currency       : " & pos.currencyCode
    Debug.Print "Quantity       : " & pos.Quantity
    Debug.Print "Market Price   : " & pos.MarketPrice

End Sub

Public Sub TestPortfolioSnapshotFactory()

    '------------------------------------------------------
    ' 1. Create the source portfolio
    '------------------------------------------------------
    Dim sourcePtf As Portfolio
    Set sourcePtf = New Portfolio

    sourcePtf.PortfolioId = "EURO_BONDS"

    '------------------------------------------------------
    ' 2. Create the source position
    '------------------------------------------------------
    Dim sourcePos As Position
    Set sourcePos = New Position

    sourcePos.PortfolioId = sourcePtf.PortfolioId
    sourcePos.InstrumentId = "OAT_FR_10Y"
    sourcePos.currencyCode = "EUR"
    sourcePos.MarketPrice = 99.5
    sourcePos.Quantity = 1000000

    sourcePtf.AddPosition sourcePos

    '------------------------------------------------------
    ' 3. Create an independent portfolio snapshot
    '------------------------------------------------------
    Dim snapshotFactory As PortfolioSnapshotFactory
    Set snapshotFactory = New PortfolioSnapshotFactory

    Dim snapshot As PortfolioSnapshot
    Set snapshot = snapshotFactory.CreateSnapshot(sourcePtf)

    '------------------------------------------------------
    ' 4. Retrieve the copied position
    '------------------------------------------------------
    Dim simulatedPos As Position
    Set simulatedPos = _
        snapshot.SimulatedPortfolio.FindPosition("OAT_FR_10Y")

    '------------------------------------------------------
    ' 5. Verify the initial equality
    '------------------------------------------------------
    Debug.Print "Source Qty before     : " & sourcePos.Quantity
    Debug.Print "Simulated Qty before  : " & simulatedPos.Quantity

    '------------------------------------------------------
    ' 6. Modify only the simulated position
    '------------------------------------------------------
    simulatedPos.Quantity = 1500000

    '------------------------------------------------------
    ' 7. Verify that the snapshot is independent
    '------------------------------------------------------
    Debug.Print "Source Qty after      : " & sourcePos.Quantity
    Debug.Print "Simulated Qty after   : " & simulatedPos.Quantity

    Debug.Print "Snapshot Id           : " & snapshot.SnapshotId
    Debug.Print "Source Portfolio Id   : " & snapshot.SourcePortfolioId
    Debug.Print "Simulated Portfolio Id: " & _
                snapshot.SimulatedPortfolio.PortfolioId

End Sub

