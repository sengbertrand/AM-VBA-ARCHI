VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutionSimulationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : ExecutionSimulationService
'
' Classification:
' Business Service / Simulation Service
'
' Domain:
' Investment / Pre-Trade Simulation
'
' Responsibility:
' Builds a simulated ExecutionReport from an Order using
' current market data.
'
' Input dependencies:
' Order
' MarketDataRepository
'
' Output:
' Simulated ExecutionReport
'
' Output consumers:
' PreTradeRiskService
' PortfolioUpdateService
'
' Mental Model:
' This service answers:
' "How would this order be executed under the current
' market assumptions?"
'==========================================================

Public Function SimulateExecution( _
                    ByRef ord As Order) As ExecutionReport

    '------------------------------------------------------
    ' Pre-condition
    '------------------------------------------------------
    If ord Is Nothing Then
        Err.Raise vbObjectError + 12001, _
                  "ExecutionSimulationService.SimulateExecution", _
                  "Order cannot be Nothing."
    End If

    '------------------------------------------------------
    ' Retrieve current market price
    '------------------------------------------------------
    Dim marketDataRepo As MarketDataRepository
    Set marketDataRepo = New MarketDataRepository

    Dim simulatedPrice As Double
    simulatedPrice = _
        marketDataRepo.GetMarketPrice(ord.InstrumentId)

    '------------------------------------------------------
    ' Build simulated execution report
    '
    ' Assumptions:
    ' - full execution
    ' - immediate execution
    ' - execution at current market price
    ' - no fees
    ' - no slippage
    '------------------------------------------------------
    Dim simulatedReport As ExecutionReport
    Set simulatedReport = New ExecutionReport

    simulatedReport.OrderId = ord.OrderId
    simulatedReport.executedQuantity = ord.Quantity
    simulatedReport.ExecutionPrice = simulatedPrice
    simulatedReport.ExecutionTime = Now
    simulatedReport.Status = ORDER_EXECUTED

    '------------------------------------------------------
    ' Return the simulated execution result
    '------------------------------------------------------
    Set SimulateExecution = simulatedReport

End Function

