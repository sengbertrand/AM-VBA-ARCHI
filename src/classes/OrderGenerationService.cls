VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderGenerationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : OrderGenerationService
'
' Classification:
' Business Service
'
' Domain:
' Investment / OMS
'
' Ownership:
' OMS owns the order generation process.
'
' Consumes:
' InvestmentDecision
'
' Produces:
' Order
'
' Responsibilities:
' Transform an investment decision into an executable order.
'
' Dependencies:
' InvestmentDecision
' Order
' Constants
'
' Mental Model:
' OrderGenerationService is the translator between PM intent
' and OMS order workflow.
'==========================================================

'-----------------------------
' Public methods
'-----------------------------

Public Function GenerateOrder(ByVal Decision As InvestmentDecision) As Order

    Dim ord As Order
    Set ord = New Order
    
    ord.OrderId = "ORD_" & Format(Now, "yyyymmdd_hhnnss")
    ord.PortfolioId = Decision.PortfolioId
    ord.instrumentId = Decision.instrumentId

    'Simplification for V1:
    'The order quantity is derived directly from target weight.
    'A future version will use OrderSizingEngine.
    
    ord.Quantity = Decision.TargetWeight * 1000000
    
    If ord.Quantity >= 0 Then
        ord.Side = SIDE_BUY
    Else
        ord.Side = SIDE_SELL
        ord.Quantity = Abs(ord.Quantity)
    End If
    
    ord.OrderDate = Date
    ord.Status = ORDER_CREATED
    
    Set GenerateOrder = ord
    
End Function

