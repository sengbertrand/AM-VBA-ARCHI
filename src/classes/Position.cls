VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : Position
'
' Classification:
' Business Object
'
' Domain:
' Investment / Portfolio / IBOR
'
' Responsibility:
' Represents the holding of one financial instrument
' within one portfolio.
'
' Ownership:
' A Portfolio owns one or more Position objects.
'
' Produced by:
' PortfolioUpdateService
'
' Consumed by:
' ValuationEngine
' Reporting
' RiskEngine
'
' Mental Model:
' A Position answers:
' "How much of this instrument does this portfolio own?"
'==========================================================

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pPortfolioId As String
Private pInstrumentId As String
Private pQuantity As Double
Private pMarketPrice As Double
Private pCurrencyCode As String

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
' Instrument identifier
'----------------------------------------------------------

Public Property Get instrumentId() As String
    instrumentId = pInstrumentId
End Property

Public Property Let instrumentId(ByVal value As String)
    pInstrumentId = value
End Property

'----------------------------------------------------------
' Position quantity
'----------------------------------------------------------

Public Property Get Quantity() As Double
    Quantity = pQuantity
End Property

Public Property Let Quantity(ByVal value As Double)
    pQuantity = value
End Property

'----------------------------------------------------------
' Latest market price
'----------------------------------------------------------

Public Property Get MarketPrice() As Double
    MarketPrice = pMarketPrice
End Property

Public Property Let MarketPrice(ByVal value As Double)
    pMarketPrice = value
End Property

'----------------------------------------------------------

'----------------------------------------------------------
' ISO currency code
' Example:
' EUR
' USD
' GBP
'----------------------------------------------------------

Public Property Get CurrencyCode() As String
    CurrencyCode = pCurrencyCode
End Property

Public Property Let CurrencyCode(ByVal value As String)
    pCurrencyCode = value
End Property
