VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InvestmentDecision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' Class : InvestmentDecision
' Purpose :
'   Represents a portfolio manager's investment decision.
'   Encapsulates the decision data through public properties.
'
' Responsibilities:
'   - Store decision attributes.
'   - Expose controlled access through Property Get/Let.
'=========================================================

Private pDecisionId As String
Private pPortfolioId As String
Private pInstrumentId As String
Private pTargetWeight As Double
Private pDecisionDate As Date
Private pRationale As String


Public Property Get DecisionId() As String
    DecisionId = pDecisionId
End Property

Public Property Let DecisionId(ByVal value As String)
    pDecisionId = value
End Property

Public Property Get PortfolioId() As String
    PortfolioId = pPortfolioId
End Property

Public Property Let PortfolioId(ByVal value As String)
    pPortfolioId = value
End Property

Public Property Get instrumentId() As String
    instrumentId = pInstrumentId
End Property

Public Property Let instrumentId(ByVal value As String)
    pInstrumentId = value
End Property

Public Property Get TargetWeight() As Double
    TargetWeight = pTargetWeight
End Property

Public Property Let TargetWeight(ByVal value As Double)
    pTargetWeight = value
End Property

Public Property Get DecisionDate() As Date
    DecisionDate = pDecisionDate
End Property

Public Property Let DecisionDate(ByVal value As Date)
    pDecisionDate = value
End Property

Public Property Get Rationale() As String
    Rationale = pRationale
End Property

Public Property Let Rationale(ByVal value As String)
    pRationale = value
End Property

