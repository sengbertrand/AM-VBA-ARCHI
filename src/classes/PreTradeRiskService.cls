VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreTradeRiskService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : PreTradeRiskService
'
' Classification:
' Business Service
'
' Domain:
' Investment / Pre-Trade Risk
'
' Responsibility:
' Evaluates the impact of a potential order before it is
' sent to the market.
'
' This service:
' - creates an independent portfolio snapshot;
' - simulates the execution;
' - calculates the execution impact;
' - applies the simulated execution impact to the snapshot;
' - evaluates the resulting portfolio concentration;
' - produces a PreTradeRiskReport.
'
' Input dependencies:
' Portfolio
' Order
'
' Uses:
' PortfolioSnapshotFactory
' ExecutionSimulationService
' InstrumentRepository
' ExecutionImpactCalculator
' PortfolioUpdateService
' RiskEngine
' ConfigRepository
'
' Produces:
' PreTradeRiskReport
'
' Consumed by:
' Calling workflow
'
' Mental Model:
' This service answers:
' "If this order were executed now, would the portfolio
' still satisfy the investment rules?"
'==========================================================

Public Function EvaluateOrder( _
                    ByVal ptf As Portfolio, _
                    ByVal ord As Order) As PreTradeRiskReport

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
    ' 3. Retrieve the instrument definition
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim currentInstr As instrument
    Set currentInstr = instrumentRepo.GetInstrument(ord.InstrumentId)

    If currentInstr Is Nothing Then
        Err.Raise vbObjectError + 13005, _
                  "PreTradeRiskService.EvaluateOrder", _
                  "Unknown instrument: " & ord.InstrumentId
    End If

    '------------------------------------------------------
    ' 4. Calculate the execution impact
    '------------------------------------------------------
    Dim impactCalc As ExecutionImpactCalculator
    Set impactCalc = New ExecutionImpactCalculator

    Dim execImpact As ExecutionImpact
    Set execImpact = impactCalc.Calculate( _
                        ord:=ord, _
                        report:=simulatedReport, _
                        currentInstr:=currentInstr)

    '------------------------------------------------------
    ' 5. Apply the simulated execution impact
    '------------------------------------------------------
    Dim updateSvc As PortfolioUpdateService
    Set updateSvc = New PortfolioUpdateService

    updateSvc.ApplyExecutionImpact _
        ptf:=snapshot.SimulatedPortfolio, _
        execImpact:=execImpact

    '------------------------------------------------------
    ' 6. Retrieve the simulated position
    '------------------------------------------------------
    Dim simulatedPos As Position
    Set simulatedPos = _
        snapshot.SimulatedPortfolio.FindPosition(ord.InstrumentId)

    If simulatedPos Is Nothing Then
        Err.Raise vbObjectError + 13004, _
                  "PreTradeRiskService.EvaluateOrder", _
                  "Simulated position not found for instrument: " & _
                  ord.InstrumentId
    End If

    '------------------------------------------------------
    ' 7. Calculate the simulated position weight
    '------------------------------------------------------
    Dim riskSvc As RiskEngine
    Set riskSvc = New RiskEngine

    Dim simulatedWeight As Double
    simulatedWeight = _
        riskSvc.CalculatePositionWeight( _
            snapshot.SimulatedPortfolio, _
            simulatedPos)

    '------------------------------------------------------
    ' 8. Retrieve the current position
    '------------------------------------------------------
    Dim currentPos As Position
    Set currentPos = ptf.FindPosition(ord.InstrumentId)

    '------------------------------------------------------
    ' 9. Calculate the current position weight
    '------------------------------------------------------
    Dim currentWeight As Double

    If currentPos Is Nothing Then

        currentWeight = 0#

    Else

        currentWeight = _
            riskSvc.CalculatePositionWeight( _
                ptf, _
                currentPos)

    End If

    '------------------------------------------------------
    ' 10. Retrieve the concentration limit
    '------------------------------------------------------
    Dim configRepo As ConfigRepository
    Set configRepo = New ConfigRepository

    Dim maximumAllowedWeight As Double
    maximumAllowedWeight = _
        configRepo.GetDoubleValue(CONFIG_MAX_POSITION_WEIGHT)

    '------------------------------------------------------
    ' 11. Build the pre-trade risk report
    '------------------------------------------------------
    Dim riskReport As PreTradeRiskReport
    Set riskReport = New PreTradeRiskReport

    riskReport.ReportId = _
        "RISK_" & Format(Now, "yyyymmdd_hhnnss")

    riskReport.PortfolioId = ptf.PortfolioId
    riskReport.OrderId = ord.OrderId
    riskReport.InstrumentId = ord.InstrumentId
    riskReport.SnapshotId = snapshot.SnapshotId

    riskReport.currentWeight = currentWeight
    riskReport.simulatedWeight = simulatedWeight
    riskReport.maximumAllowedWeight = maximumAllowedWeight
    riskReport.DecisionTime = Now

    '------------------------------------------------------
    ' 12. Apply the concentration rule
    '------------------------------------------------------
    If simulatedWeight <= maximumAllowedWeight Then

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
    ' 13. Return the completed report
    '------------------------------------------------------
    Set EvaluateOrder = riskReport

End Function
