VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Factory : PositionFactory
'
' Classification:
' Factory
'
' Domain:
' Investment / Portfolio / IBOR
'
' Responsibility:
' Creates Position objects from business context.
'
' Producer:
' PortfolioUpdateService
'
' Consumer:
' Portfolio
'
' Mental Model:
' This factory answers:
' "How should a new Position be initialized?"
'==========================================================

Public Function CreateFromOrder(ByRef ptf As Portfolio, _
                                ByRef ord As Order) As Position

    If ptf Is Nothing Then
        Err.Raise vbObjectError + 4001, _
                  "PositionFactory.CreateFromOrder", _
                  "Portfolio cannot be Nothing."
    End If

    If ord Is Nothing Then
        Err.Raise vbObjectError + 4002, _
                  "PositionFactory.CreateFromOrder", _
                  "Order cannot be Nothing."
    End If

    Dim pos As Position
    Set pos = New Position

    pos.PortfolioId = ptf.PortfolioId
    pos.instrumentId = ord.instrumentId
    pos.Quantity = 0
    pos.MarketPrice = 0

    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    Dim foundInstrument As Instrument
    Set foundInstrument = instrumentRepo.GetInstrument(ord.instrumentId)

    pos.CurrencyCode = foundInstrument.CurrencyCode

    Set CreateFromOrder = pos

End Function
