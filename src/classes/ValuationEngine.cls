VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuationEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========================================================
' Business Service : ValuationEngine
'
' Classification:
' Business Service / Calculation Engine
'
' Domain:
' Investment / Portfolio Valuation
'
' Responsibility:
' Calculates the market value of a Position according
' to the characteristics of its Instrument.
'
' Input dependencies:
' Position
' Instrument
'
' Output:
' Market Value
'
' Output consumers:
' Reporting
' RiskEngine
' PerformanceEngine
'
' Mental Model:
' This engine answers:
' "What is this position worth?"
'==========================================================

Public Function CalculateMarketValue(ByRef pos As Position, _
                                     ByRef foundInstrument As Instrument) As Double

    '------------------------------------------------------
    ' Pre-conditions
    '------------------------------------------------------
    If pos Is Nothing Then
        Err.Raise vbObjectError + 6001, _
                  "ValuationEngine.CalculateMarketValue", _
                  "Position cannot be Nothing."
    End If

    If foundInstrument Is Nothing Then
        Err.Raise vbObjectError + 6002, _
                  "ValuationEngine.CalculateMarketValue", _
                  "Instrument cannot be Nothing."
    End If

    '------------------------------------------------------
    ' Apply the valuation convention by asset class
    '------------------------------------------------------
    If foundInstrument.AssetClass = ASSET_CLASS_BOND Then

        CalculateMarketValue = _
            pos.Quantity * pos.MarketPrice / 100

    ElseIf foundInstrument.AssetClass = ASSET_CLASS_EQUITY Then

        CalculateMarketValue = _
            pos.Quantity * pos.MarketPrice

    Else
        Err.Raise vbObjectError + 6003, _
                  "ValuationEngine.CalculateMarketValue", _
                  "Unsupported asset class: " & _
                  foundInstrument.AssetClass
    End If

End Function
