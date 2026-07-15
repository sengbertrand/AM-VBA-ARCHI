VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreTradeRiskService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function EvaluateOrder( _
                    ByRef ptf As Portfolio, _
                    ByRef ord As Order) As PreTradeRiskReport

    '------------------------------------------------------
    ' Pre-conditions
    '------------------------------------------------------
    If ptf Is Nothing Then
        Err.Raise vbObjectError + 13001, _
                  "PreTradeRiskService.EvaluateOrder", _
                  "Portfolio cannot be Nothing."
    End If

    If ord Is Nothing Then
        Err.Raise vbObjectError + 13002, _
                  "PreTradeRiskService.EvaluateOrder", _
                  "Order cannot be Nothing."
    End If

    If ord.PortfolioId <> ptf.PortfolioId Then
        Err.Raise vbObjectError + 13003, _
                  "PreTradeRiskService.EvaluateOrder", _
                  "Order portfolio does not match target portfolio."
    End If

    '------------------------------------------------------
    ' 1. Create an independent portfolio snapshot
    '------------------------------------------------------
    Dim snapshotFactory As PortfolioSnapshotFactory
    Set snapshotFactory = New PortfolioSnapshotFactory

    Dim snapshot As PortfolioSnapshot
    Set snapshot = snapshotFactory.CreateSnapshot(ptf)

    '------------------------------------------------------
    ' 2. Simulate order execution
    '------------------------------------------------------
    Dim simulationSvc As ExecutionSimulationService
    Set simulationSvc = New ExecutionSimulationService

    Dim simulatedReport As ExecutionReport
    Set simulatedReport = simulationSvc.SimulateExecution(ord)

    '------------------------------------------------------
    ' 3. Apply the simulated execution to the snapshot
    '------------------------------------------------------
    Dim updateSvc As PortfolioUpdateService
    Set updateSvc = New PortfolioUpdateService

    updateSvc.ApplyExecution _
        snapshot.SimulatedPortfolio, _
        ord, _
        simulatedReport

    '------------------------------------------------------
    ' 4. Retrieve the simulated position
    '------------------------------------------------------
    Dim simulatedPos As Position
    Set simulatedPos = _
        snapshot.SimulatedPortfolio.FindPosition(ord.instrumentId)

    If simulatedPos Is Nothing Then
        Err.Raise vbObjectError + 13004, _
                  "PreTradeRiskService.EvaluateOrder", _
                  "Simulated position not found for instrument: " & _
                  ord.instrumentId
    End If

    '------------------------------------------------------
    ' 5. Calculate the simulated position weight
    '------------------------------------------------------
    Dim riskEngin As RiskEngine
    Set riskEngin = New RiskEngine

    Dim SimulatedWeight As Double
    SimulatedWeight = _
        riskEngin.CalculatePositionWeight( _
            snapshot.SimulatedPortfolio, _
            simulatedPos)

    '------------------------------------------------------
    ' 6. Retrieve the current position
    '------------------------------------------------------
    Dim currentPos As Position
    Set currentPos = ptf.FindPosition(ord.instrumentId)

    '------------------------------------------------------
    ' 7. Calculate the current position weight
    '------------------------------------------------------
    Dim CurrentWeight As Double

    If currentPos Is Nothing Then

        CurrentWeight = 0

    Else

        CurrentWeight = _
            riskEngin.CalculatePositionWeight( _
                ptf, _
                currentPos)

    End If

    '------------------------------------------------------
    ' 8. Define the maximum allowed weight
    '-----------------------------------------------------
    
    Dim configRepo As ConfigRepository
    Set configRepo = New ConfigRepository
    
    Dim MaximumAllowedWeight As Double
    MaximumAllowedWeight = configRepo.GetDoubleValue( _
            CONFIG_MAX_POSITION_WEIGHT)

    '------------------------------------------------------
    ' 9. Build the pre-trade risk report
    '------------------------------------------------------
    Dim riskReport As PreTradeRiskReport
    Set riskReport = New PreTradeRiskReport

    riskReport.ReportId = _
        "RISK_" & Format(Now, "yyyymmdd_hhnnss")

    riskReport.PortfolioId = ptf.PortfolioId
    riskReport.OrderId = ord.OrderId
    riskReport.instrumentId = ord.instrumentId
    riskReport.SnapshotId = snapshot.SnapshotId

    riskReport.CurrentWeight = CurrentWeight
    riskReport.SimulatedWeight = SimulatedWeight
    riskReport.MaximumAllowedWeight = MaximumAllowedWeight
    riskReport.DecisionTime = Now

    '------------------------------------------------------
    ' 10. Apply the temporary concentration rule
    '------------------------------------------------------
    If SimulatedWeight <= MaximumAllowedWeight Then

        riskReport.Approved = True
        riskReport.Decision = RISK_APPROVED
        riskReport.Reason = _
            "The simulated position remains within the " & _
            "maximum concentration limit."

    Else

        riskReport.Approved = False
        riskReport.Decision = RISK_REJECTED
        riskReport.Reason = _
            "The simulated position exceeds the " & _
            "maximum concentration limit."

    End If

    '------------------------------------------------------
    ' 11. Return the completed report
    '------------------------------------------------------
    Set EvaluateOrder = riskReport

End Function

