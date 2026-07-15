VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PortfolioValuationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : PortfolioValuationService
'
' Classification:
' Business Service / Aggregation Service
'
' Domain:
' Investment / Portfolio Valuation
'
' Responsibility:
' Calculates the total market value of a Portfolio by
' valuing and aggregating each of its Positions.
'
' Input dependencies:
' Portfolio
' InstrumentRepository
' ValuationEngine
'
' Output:
' Portfolio Market Value
'
' Output consumers:
' Reporting
' RiskEngine
' PerformanceEngine
'
' Mental Model:
' This service answers:
' "What is the total value of this portfolio?"
'==========================================================

Public Function CalculatePortfolioMarketValue( _
                    ByRef ptf As Portfolio) As Double

    '------------------------------------------------------
    ' Pre-condition
    '------------------------------------------------------
    If ptf Is Nothing Then
        Err.Raise vbObjectError + 7001, _
                  "PortfolioValuationService." & _
                  "CalculatePortfolioMarketValue", _
                  "Portfolio cannot be Nothing."
    End If

    '------------------------------------------------------
    ' Create required service dependencies
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim valEngine As ValuationEngine
    Set valEngine = New ValuationEngine

    '------------------------------------------------------
    ' Aggregate the value of all portfolio positions
    '------------------------------------------------------
    Dim totalMarketValue As Double
    totalMarketValue = 0

    Dim pos As Position
    Dim foundInstrument As Instrument
    Dim positionMarketValue As Double

    For Each pos In ptf.Positions

        Set foundInstrument = _
            instrumentRepo.GetInstrument(pos.instrumentId)

        positionMarketValue = _
            valEngine.CalculateMarketValue(pos, foundInstrument)

        totalMarketValue = _
            totalMarketValue + positionMarketValue

    Next pos

    CalculatePortfolioMarketValue = totalMarketValue

End Function

