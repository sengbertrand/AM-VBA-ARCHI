Attribute VB_Name = "TestPortfolioWorkflow"
Option Explicit

Public Sub TestPortfolioUpdateService()

    '------------------------------------------------------
    ' Create initial portfolio
    '------------------------------------------------------
    
    Dim ptf As Portfolio
    Set ptf = New Portfolio
    
    ptf.PortfolioId = "EURO_BONDS"
    
    '------------------------------------------------------
    ' Create initial position
    '------------------------------------------------------
    
    Dim pos As Position
    Set pos = New Position
    
    pos.PortfolioId = "EURO_BONDS"
    pos.instrumentId = "OAT_FR_10Y"
    pos.Quantity = 1000000
    pos.MarketPrice = 99.5
    pos.CurrencyCode = "EUR"
    
    ptf.AddPosition pos
    
    Debug.Print "Initial Qty : " & pos.Quantity
    
    '------------------------------------------------------
    ' Create executed order
    '------------------------------------------------------
    
    Dim ord As Order
    Set ord = New Order
    
    ord.OrderId = "ORD_002"
    ord.PortfolioId = "EURO_BONDS"
    ord.instrumentId = "OAT_FR_10Y"
    ord.Side = SIDE_BUY
    ord.Quantity = 300000
    ord.Status = ORDER_CREATED
    
    '------------------------------------------------------
    ' Execute order
    '------------------------------------------------------
    
    Dim execService As ExecutionService
    Set execService = New ExecutionService
    
    Dim Report As ExecutionReport
    Set Report = execService.ExecuteOrder(ord)
    
    '------------------------------------------------------
    ' Apply execution to portfolio
    '------------------------------------------------------
    
    Dim updateService As PortfolioUpdateService
    Set updateService = New PortfolioUpdateService
    
    updateService.ApplyExecution ptf, ord, Report
    
    Debug.Print "Final Qty   : " & pos.Quantity
    Debug.Print "Final Price : " & pos.MarketPrice
    
End Sub

Public Sub TestEndToEndWorkflow()

    '------------------------------------------------------
    ' 1. Create initial portfolio
    '------------------------------------------------------
    
    Dim ptf As Portfolio
    Set ptf = New Portfolio
    
    ptf.PortfolioId = "EURO_BONDS"
    
    '------------------------------------------------------
    ' 2. Create initial position
    '------------------------------------------------------
    
    Dim pos As Position
    Set pos = New Position
    
    pos.PortfolioId = "EURO_BONDS"
    pos.instrumentId = "OAT_FR_10Y"
    pos.Quantity = 1000000
    pos.MarketPrice = 99.5
    pos.CurrencyCode = "EUR"
    
    ptf.AddPosition pos
    
    Debug.Print "Initial Qty : " & pos.Quantity
        
    '------------------------------------------------------
    ' 3. Create investment decision
    '------------------------------------------------------
    
    Dim Decision As InvestmentDecision
    Set Decision = New InvestmentDecision
    
    Decision.DecisionId = "DEC_002"
    Decision.PortfolioId = "EURO_BONDS"
    Decision.instrumentId = "OAT_FR_10Y"
    Decision.TargetWeight = 0.15
    Decision.DecisionDate = Date
    Decision.Rationale = "Increase exposure to French government bonds"
    
    '------------------------------------------------------
    ' 4. Generate order from decision
    '------------------------------------------------------
    
    Dim orderService As OrderGenerationService
    Set orderService = New OrderGenerationService
    
    Dim ord As Order
    Set ord = orderService.GenerateOrder(Decision)
    
    Debug.Print "Generated Order : " & ord.OrderId
    Debug.Print "Order Qty           : " & ord.Quantity
    Debug.Print "Order Status       : " & ord.Status
    
        
    '------------------------------------------------------
    ' 5. Execute order
    '------------------------------------------------------
    
    Dim execService As ExecutionService
    Set execService = New ExecutionService
    
    Dim Report As ExecutionReport
    Set Report = execService.ExecuteOrder(ord)
    
    Debug.Print "Execution Status : " & Report.Status
    Debug.Print "Execution Price    : " & Report.ExecutionPrice
    
        
    '------------------------------------------------------
    ' 6. Apply execution to portfolio
    '------------------------------------------------------
    
    Dim updateService As PortfolioUpdateService
    Set updateService = New PortfolioUpdateService
    
    updateService.ApplyExecution ptf, ord, Report
    
    '------------------------------------------------------
    ' 7. Verify final portfolio state
    '------------------------------------------------------
    Debug.Print "Final Qty   : " & pos.Quantity
    Debug.Print "Final Price : " & pos.MarketPrice

End Sub

Public Sub TestNewPositionCreation()
    
    Dim ptf As Portfolio
    Set ptf = New Portfolio
    
    ptf.PortfolioId = "EURO_BONDS"
    
    Dim ord As Order
    Set ord = New Order
    
    ord.OrderId = "ORD_003"
    ord.PortfolioId = ptf.PortfolioId
    ord.instrumentId = "BUND_DE_10Y"
    ord.Side = SIDE_BUY
    ord.Quantity = 500000
    ord.Status = ORDER_CREATED
    
    Dim execService As ExecutionService
    Set execService = New ExecutionService
    
    Dim execReport As ExecutionReport
    Set execReport = execService.ExecuteOrder(ord)
    
    Dim posUpdateService As PortfolioUpdateService
    Set posUpdateService = New PortfolioUpdateService
    
    posUpdateService.ApplyExecution ptf, ord, execReport
    
    Dim pos As Position
    Set pos = ptf.FindPosition("BUND_DE_10Y")
    
    Debug.Print "Position Count : " & ptf.PositionCount
    Debug.Print "Instrument         : " & pos.instrumentId
    Debug.Print "Currency           : " & pos.CurrencyCode
    Debug.Print "Quantity             : " & pos.Quantity
    Debug.Print "Market Price        : " & pos.MarketPrice

End Sub

Public Sub TestPortfolioSnapshotFactory()
    
    '----------------------
    ' Create source portfolio
    '---------------------
    
    Dim sourcePtf As Portfolio
    Set sourcePtf = New Portfolio
    
    sourcePtf.PortfolioId = "EURO_BONDS"
    
    '----------------------------------
    ' Create source position
    '--------------------------
    
    Dim sourcePos As Position
    Set sourcePos = New Position
    
    sourcePos.PortfolioId = sourcePtf.PortfolioId
    sourcePos.instrumentId = "OAT_FR_10Y"
    sourcePos.CurrencyCode = "EUR"
    sourcePos.MarketPrice = 99.5
    sourcePos.Quantity = 1000000
    
    '-------------------------------
    ' Add sourcePos to sourcePtf
    '----------------------------
    
    sourcePtf.AddPosition sourcePos
    
    
    '-----------------------------
    ' Create independent snapshot
    '-------------------------------
    
    Dim snapshotFactory As PortfolioSnapshotFactory
    Set snapshotFactory = New PortfolioSnapshotFactory
    
    Dim snapshot As PortfolioSnapshot
    Set snapshot = snapshotFactory.CreateSnapshot(sourcePtf)
    
    '---------------------------------------
    ' Retrieve copied position from simulated portfolio
    '-----------------------------------------------
    Dim simulatedPos As Position
    Set simulatedPos = snapshot.SimulatedPortfolio.FindPosition("OAT_FR_10Y")
    
    '-------------------------
    ' Verify initial equality
    '---------------------
    
    Debug.Print "Source Qty before : " & sourcePos.Quantity
    Debug.Print "Simulated Qty before : " & simulatedPos.Quantity
    
    '------------------------------------
    ' Modify only the simulated position
    '---------------------------------
    
    simulatedPos.Quantity = 1500000
    
    '----------------------------------
    ' Verify object independance
    '----------------------------
    
    Debug.Print "Source Qty after : " & sourcePos.Quantity
    Debug.Print "Simulated Qty after : " & simulatedPos.Quantity
    
    Debug.Print "Snapshot Id                : " & snapshot.SnapshotId
    Debug.Print "Source Portfolio Id     : " & snapshot.SourcePortfolioId
    Debug.Print "Simulated Portfolio Id: " & snapshot.SimulatedPortfolio.PortfolioId
    
End Sub
