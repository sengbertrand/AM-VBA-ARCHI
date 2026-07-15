Attribute VB_Name = "TestOrderWorkFlow"
Option Explicit


Public Sub TestOrder()

    '------------------------------------------------------
    ' Test:
    ' Create an Order business object.
    '------------------------------------------------------
    
    Dim ord As Order
    Set ord = New Order
    
    '-----------------------------
    ' Populate business data
    '-----------------------------
    
    ord.OrderId = "ORD_001"
    ord.PortfolioId = "EURO_BONDS"
    ord.instrumentId = "OAT_FR_10Y"
    
    ord.Side = SIDE_BUY
    ord.Quantity = 3000000
    
    ord.OrderDate = Date
    ord.Status = ORDER_CREATED
    
    '-----------------------------
    ' Verify object content
    '-----------------------------
    
    Debug.Print "Order Id                : "; ord.OrderId
    Debug.Print "Portfolio               : "; ord.PortfolioId
    Debug.Print "Instrument            : "; ord.instrumentId
    Debug.Print "Side                      : "; ord.Side
    Debug.Print "Quantity               : "; ord.Quantity
    Debug.Print "Status                   : "; ord.Status
    
    '-----------------------------
    ' Log test result
    '-----------------------------
    
    LogInfo "Order created : " & ord.OrderId
    
    
End Sub

Public Sub TestOrderGenerationService()

    Dim Decision As InvestmentDecision
    Set Decision = New InvestmentDecision
    
    Decision.DecisionId = "DEC_001"
    Decision.PortfolioId = "EURO_BONDS"
    Decision.instrumentId = "OAT_FR_10Y"
    Decision.TargetWeight = 0.15
    Decision.DecisionDate = Date
    Decision.Rationale = "Increase exposure to French government bonds"
    
    Dim orderService As OrderGenerationService
    Set orderService = New OrderGenerationService
    
    Dim ord As Order
    Set ord = orderService.GenerateOrder(Decision)
    
    Debug.Print "Generated Order Id : "; ord.OrderId
    Debug.Print "Portfolio                 : "; ord.PortfolioId
    Debug.Print "Instrument             : "; ord.instrumentId
    Debug.Print "Side                        : "; ord.Side
    Debug.Print "Quantity                 : "; ord.Quantity
    Debug.Print "Status                     : "; ord.Status
    
    LogInfo "Order generated : " & ord.OrderId & _
            " | " & ord.Side & _
            " | " & ord.instrumentId & _
            " | Qty =" & ord.Quantity
            
    Dim execService As ExecutionService
    Set execService = New ExecutionService
    
    Dim Report As ExecutionReport
    Set Report = execService.ExecuteOrder(ord)
    
    Debug.Print "Order Status after execution : " & ord.Status
    Debug.Print "Execution Report OrderId : " & Report.OrderId
    Debug.Print "Executed Quantity : " & Report.ExecutedQuantity
    Debug.Print "Execution Price : " & Report.ExecutionPrice
    Debug.Print "Execution Time : " & Report.ExecutionTime
    Debug.Print "Execution Status : " & Report.Status
    
End Sub

Public Sub TestExecutionSimulationService()

    '-----------
    'Create Order
    '--------------
    
    Dim ord As Order
    Set ord = New Order
    
    
    ord.OrderId = "ORD_SIM_001"
    ord.PortfolioId = "EURO_BONDS"
    ord.instrumentId = "BUND_DE_10Y"
    ord.Side = SIDE_BUY
    ord.Quantity = 500000
    ord.Status = ORDER_CREATED
    
    '------------------
    ' Simulate execution
    '----------------------
    
    Dim simulationSvc As ExecutionSimulationService
    Set simulationSvc = New ExecutionSimulationService
    
    Dim simulatedReport As ExecutionReport
    Set simulatedReport = simulationSvc.SimulateExecution(ord)
    
    '---------
    ' Verify simulated report
    '---------------------
    
    Debug.Print "Order  Id               : " & simulatedReport.OrderId
    Debug.Print "Executed Quantity : " & simulatedReport.ExecutedQuantity
    Debug.Print "Executed Price       : " & simulatedReport.ExecutionPrice
    Debug.Print "Execution TIme      : " & simulatedReport.ExecutionTime
    Debug.Print "Execution Status    : " & simulatedReport.Status
    
    '-------------
    ' Verify that the real Order was not executed
    '--------------
    
    Debug.Print "Real order status : " & ord.Status
    

End Sub

