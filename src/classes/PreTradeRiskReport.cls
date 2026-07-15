VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreTradeRiskReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : PreTradeRiskReport
'
' Classification:
' Business Object / Decision Report
'
' Domain:
' Investment / Pre-Trade Risk
'
' Responsibility:
' Represents the result of a pre-trade risk assessment
' performed on an Order against a simulated Portfolio.
'
' Ownership:
' PreTradeRiskService owns the report lifecycle.
'
' Produced by:
' PreTradeRiskService
'
' Consumed by:
' Order workflow
' ExecutionService
' Reporting
' Audit
'
' Mental Model:
' A PreTradeRiskReport answers:
' "Can this order be executed, and why?"
'==========================================================

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pReportId As String
Private pPortfolioId As String
Private pOrderId As String
Private pInstrumentId As String
Private pSnapshotId As String

Private pApproved As Boolean
Private pDecision As String
Private pReason As String

Private pCurrentWeight As Double
Private pSimulatedWeight As Double
Private pMaximumAllowedWeight As Double

Private pDecisionTime As Date

'----------------------------------------------------------
' Report identifier
'----------------------------------------------------------

Public Property Get ReportId() As String
    ReportId = pReportId
End Property

Public Property Let ReportId(ByVal value As String)
    pReportId = value
End Property

'----------------------------------------------------------
' Portfolio identifier
'----------------------------------------------------------

Public Property Get PortfolioId() As String
    PortfolioId = pPortfolioId
End Property

Public Property Let PortfolioId(ByVal value As String)
    pPortfolioId = value
End Property

'----------------------------------------------------------
' Order identifier
'----------------------------------------------------------

Public Property Get OrderId() As String
    OrderId = pOrderId
End Property

Public Property Let OrderId(ByVal value As String)
    pOrderId = value
End Property

'----------------------------------------------------------
' Instrument identifier
'----------------------------------------------------------

Public Property Get instrumentId() As String
    instrumentId = pInstrumentId
End Property

Public Property Let instrumentId(ByVal value As String)
    pInstrumentId = value
End Property

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
' Approval result
'----------------------------------------------------------

Public Property Get Approved() As Boolean
    Approved = pApproved
End Property

Public Property Let Approved(ByVal value As Boolean)
    pApproved = value
End Property

'----------------------------------------------------------
' Decision label
' Example: APPROVED or REJECTED
'----------------------------------------------------------

Public Property Get Decision() As String
    Decision = pDecision
End Property

Public Property Let Decision(ByVal value As String)
    pDecision = value
End Property

'----------------------------------------------------------
' Decision explanation
'----------------------------------------------------------

Public Property Get Reason() As String
    Reason = pReason
End Property

Public Property Let Reason(ByVal value As String)
    pReason = value
End Property

'----------------------------------------------------------
' Current position weight
'----------------------------------------------------------

Public Property Get CurrentWeight() As Double
    CurrentWeight = pCurrentWeight
End Property

Public Property Let CurrentWeight(ByVal value As Double)
    pCurrentWeight = value
End Property

'----------------------------------------------------------
' Simulated position weight after the order
'----------------------------------------------------------

Public Property Get SimulatedWeight() As Double
    SimulatedWeight = pSimulatedWeight
End Property

Public Property Let SimulatedWeight(ByVal value As Double)
    pSimulatedWeight = value
End Property

'----------------------------------------------------------
' Maximum permitted position weight
'----------------------------------------------------------

Public Property Get MaximumAllowedWeight() As Double
    MaximumAllowedWeight = pMaximumAllowedWeight
End Property

Public Property Let MaximumAllowedWeight(ByVal value As Double)
    pMaximumAllowedWeight = value
End Property

'----------------------------------------------------------
' Decision timestamp
'----------------------------------------------------------

Public Property Get DecisionTime() As Date
    DecisionTime = pDecisionTime
End Property

Public Property Let DecisionTime(ByVal value As Date)
    pDecisionTime = value
End Property
