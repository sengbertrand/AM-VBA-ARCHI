VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : Order
'
' Classification:
' Business Object
'
' Domain:
' Investment / OMS
'
' Ownership:
' OMS owns the order lifecycle.
'
' Produced by:
' OrderGenerationService
'
' Consumed by:
' ComplianceEngine
' PreTradeRiskEngine
' ExecutionService
'
' Responsibilities:
' Store an executable order before market execution.
'
' Dependencies:
' None. This class must not access Excel directly.
'
' Mental Model:
' Order is the operational translation of an investment decision.
'==========================================================

'-----------------------------
' Private fields
'-----------------------------

Private pOrderId As String
Private pPortfolioId As String
Private pInstrumentId As String
Private pSide As String
Private pQuantity As Double
Private pOrderDate As Date
Private pStatus As String

'-----------------------------
' Order identifier
'-----------------------------

Public Property Get OrderId() As String

    OrderId = pOrderId
    
End Property

Public Property Let OrderId(ByVal value As String)

    pOrderId = value
    
End Property

'------------------------------
' Portfolio identifier
'------------------------------

Public Property Get PortfolioId() As String

    PortfolioId = pPortfolioId
    
End Property

Public Property Let PortfolioId(ByVal value As String)

    pPortfolioId = value
    
End Property

'-------------------------------
' Instrument identifier
'-------------------------------

Public Property Get instrumentId() As String

    instrumentId = pInstrumentId
    
End Property

Public Property Let instrumentId(ByVal value As String)

    pInstrumentId = value
    
End Property

'-----------------------------
' Order side: BUY or SELL
'-----------------------------

Public Property Get Side() As String

    Side = pSide
    
End Property

Public Property Let Side(ByVal value As String)

    pSide = value
    
End Property

'-----------------------------
' Order quantity
'-----------------------------

Public Property Get Quantity() As Double

    Quantity = pQuantity
    
End Property

Public Property Let Quantity(ByVal value As Double)

    pQuantity = value
    
End Property

'-----------------------------
' Order Date
'-----------------------------

Public Property Get OrderDate() As Date

    OrderDate = pOrderDate
    
End Property

Public Property Let OrderDate(ByVal value As Date)

    pOrderDate = value
    
End Property

'-----------------------------
' Order lifecycle status
'-----------------------------

Public Property Get Status() As String

    Status = pStatus
    
End Property

Public Property Let Status(ByVal value As String)

    pStatus = value
    
End Property
