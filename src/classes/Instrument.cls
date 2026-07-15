VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Instrument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : Instrument
'
' Classification:
' Business Object
'
' Domain:
' Investment / Security Master
'
' Responsibility:
' Represents the reference data of one financial instrument.
'
' Ownership:
' Instrument reference data is owned by the Security Master.
'
' Produced by:
' InstrumentRepository
'
' Consumed by:
' PositionFactory
' ValuationEngine
' Reporting
' RiskEngine
'
' Mental Model:
' An Instrument answers:
' "What financial security are we dealing with?"
'==========================================================

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pInstrumentId As String
Private pInstrumentName As String
Private pAssetClass As String
Private pCurrencyCode As String

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
' Instrument name
'----------------------------------------------------------

Public Property Get InstrumentName() As String
    InstrumentName = pInstrumentName
End Property

Public Property Let InstrumentName(ByVal value As String)
    pInstrumentName = value
End Property

'----------------------------------------------------------
' Asset class
'----------------------------------------------------------

Public Property Get AssetClass() As String
    AssetClass = pAssetClass
End Property

Public Property Let AssetClass(ByVal value As String)
    pAssetClass = value
End Property

'----------------------------------------------------------
' ISO currency code
'----------------------------------------------------------

Public Property Get CurrencyCode() As String
    CurrencyCode = pCurrencyCode
End Property

Public Property Let CurrencyCode(ByVal value As String)
    pCurrencyCode = value
End Property
