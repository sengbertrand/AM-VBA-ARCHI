VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PortfolioSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : PortfolioSnapshot
'
' Classification:
' Business Object / Simulation State
'
' Domain:
' Investment / Portfolio / Pre-Trade Risk
'
' Responsibility:
' Represents an independent copy of a Portfolio at a
' specific point in time for simulation purposes.
'
' Ownership:
' PreTradeRiskService owns the simulation lifecycle.
'
' Produced by:
' PortfolioSnapshotFactory
'
' Consumed by:
' PreTradeRiskService
' RiskEngine
' PortfolioValuationService
'
' Mental Model:
' A PortfolioSnapshot answers:
' "What would this portfolio look like under a simulated
' scenario without modifying the real portfolio?"
'==========================================================

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pSnapshotId As String
Private pSnapshotTime As Date
Private pSourcePortfolioId As String
Private pSimulatedPortfolio As Portfolio

'----------------------------------------------------------
' Snapshot identifier
'----------------------------------------------------------

Public Property Get SnapshotId() As String
    SnapshotId = pSnapshotId
End Property

Public Property Let SnapshotId(ByVal value As String)
    pSnapshotId = value
End Property

'----------------------------------------------------------
' Snapshot creation timestamp
'----------------------------------------------------------

Public Property Get SnapshotTime() As Date
    SnapshotTime = pSnapshotTime
End Property

Public Property Let SnapshotTime(ByVal value As Date)
    pSnapshotTime = value
End Property

'----------------------------------------------------------
' Identifier of the original portfolio
'----------------------------------------------------------

Public Property Get SourcePortfolioId() As String
    SourcePortfolioId = pSourcePortfolioId
End Property

Public Property Let SourcePortfolioId(ByVal value As String)
    pSourcePortfolioId = value
End Property

'----------------------------------------------------------
' Independent simulated portfolio
'----------------------------------------------------------

Public Property Get SimulatedPortfolio() As Portfolio
    Set SimulatedPortfolio = pSimulatedPortfolio
End Property

Public Property Set SimulatedPortfolio(ByVal value As Portfolio)

    If value Is Nothing Then
        Err.Raise vbObjectError + 10001, _
                  "PortfolioSnapshot.SimulatedPortfolio", _
                  "Simulated portfolio cannot be Nothing."
    End If

    Set pSimulatedPortfolio = value

End Property

