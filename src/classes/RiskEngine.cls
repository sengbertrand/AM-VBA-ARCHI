VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiskEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : RiskEngine
'
' Classification:
' Business Service / Risk Calculation Engine
'
' Domain:
' Investment / Portfolio Risk
'
' Responsibility:
' Calculates portfolio risk indicators from positions,
' instruments and portfolio valuations.
'
' Input dependencies:
' Portfolio
' Position
' InstrumentRepository
' ValuationEngine
' PortfolioValuationService
'
' Output:
' Portfolio risk metrics
'
' Output consumers:
' PreTradeRiskService
' Reporting
' Portfolio Management
'
' Mental Model:
' This engine answers:
' "What risks are created by the portfolio holdings?"
'==========================================================

Public Function CalculatePositionWeight( _
                    ByRef ptf As Portfolio, _
                    ByRef pos As Position) As Double

    '------------------------------------------------------
    ' Pre-conditions
    '------------------------------------------------------
    If ptf Is Nothing Then
        Err.Raise vbObjectError + 8001, _
                  "RiskEngine.CalculatePositionWeight", _
                  "Portfolio cannot be Nothing."
    End If

    If pos Is Nothing Then
        Err.Raise vbObjectError + 8002, _
                  "RiskEngine.CalculatePositionWeight", _
                  "Position cannot be Nothing."
    End If

    '------------------------------------------------------
    ' Retrieve the instrument reference data
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim foundInstrument As Instrument
    Set foundInstrument = _
        instrumentRepo.GetInstrument(pos.instrumentId)

    '------------------------------------------------------
    ' Calculate the position market value
    '------------------------------------------------------
    Dim valEngine As ValuationEngine
    Set valEngine = New ValuationEngine

    Dim positionMarketValue As Double
    positionMarketValue = _
        valEngine.CalculateMarketValue(pos, foundInstrument)

    '------------------------------------------------------
    ' Calculate the total portfolio market value
    '------------------------------------------------------
    Dim portfolioValuation As PortfolioValuationService
    Set portfolioValuation = New PortfolioValuationService

    Dim totalMarketValue As Double
    totalMarketValue = _
        portfolioValuation.CalculatePortfolioMarketValue(ptf)

    '------------------------------------------------------
    ' Protect against division by zero
    '------------------------------------------------------
    If totalMarketValue = 0 Then
        Err.Raise vbObjectError + 8003, _
                  "RiskEngine.CalculatePositionWeight", _
                  "Portfolio market value cannot be zero."
    End If

    '------------------------------------------------------
    ' Calculate position concentration
    '------------------------------------------------------
    CalculatePositionWeight = _
        positionMarketValue / totalMarketValue

End Function

