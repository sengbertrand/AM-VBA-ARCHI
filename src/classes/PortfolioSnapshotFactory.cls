VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PortfolioSnapshotFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Factory : PortfolioSnapshotFactory
'
' Classification:
' Factory / Deep Copy Builder
'
' Domain:
' Investment / Portfolio / Pre-Trade Risk
'
' Responsibility:
' Creates an independent PortfolioSnapshot from a
' source Portfolio for simulation purposes.
'
' Input dependencies:
' Portfolio
' Position
'
' Output:
' PortfolioSnapshot
'
' Output consumers:
' PreTradeRiskService
' RiskEngine
' PortfolioValuationService
'
' Mental Model:
' This factory answers:
' "How do we create an independent simulation state
' without modifying the real portfolio?"
'==========================================================

Public Function CreateSnapshot(ByRef sourcePtf As Portfolio) _
                               As PortfolioSnapshot

    '------------------------------------------------------
    ' Pre-condition
    '------------------------------------------------------
    If sourcePtf Is Nothing Then
        Err.Raise vbObjectError + 11001, _
                  "PortfolioSnapshotFactory.CreateSnapshot", _
                  "Source portfolio cannot be Nothing."
    End If

    '------------------------------------------------------
    ' Create the independent simulated portfolio
    '------------------------------------------------------
    Dim simulatedPtf As Portfolio
    Set simulatedPtf = New Portfolio

    simulatedPtf.PortfolioId = sourcePtf.PortfolioId

    '------------------------------------------------------
    ' Deep-copy every position
    '------------------------------------------------------
    Dim sourcePos As Position
    Dim copiedPos As Position

    For Each sourcePos In sourcePtf.Positions

        Set copiedPos = New Position

        copiedPos.PortfolioId = sourcePos.PortfolioId
        copiedPos.instrumentId = sourcePos.instrumentId
        copiedPos.Quantity = sourcePos.Quantity
        copiedPos.MarketPrice = sourcePos.MarketPrice
        copiedPos.CurrencyCode = sourcePos.CurrencyCode

        simulatedPtf.AddPosition copiedPos

    Next sourcePos

    '------------------------------------------------------
    ' Create the snapshot wrapper
    '------------------------------------------------------
    Dim snapshot As PortfolioSnapshot
    Set snapshot = New PortfolioSnapshot

    snapshot.SnapshotId = _
        "SNAP_" & Format(Now, "yyyymmdd_hhnnss")

    snapshot.SnapshotTime = Now
    snapshot.SourcePortfolioId = sourcePtf.PortfolioId

    Set snapshot.SimulatedPortfolio = simulatedPtf

    Set CreateSnapshot = snapshot

End Function

