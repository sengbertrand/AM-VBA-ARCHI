VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Portfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : Portfolio
'
' Classification:
' Business Object
'
' Domain:
' Investment / Portfolio / IBOR
'
' Responsibility:
' Stores the security positions and cash positions
' belonging to the same investment portfolio.
'
' Ownership:
' The Portfolio owns its Position and CashPosition objects.
'
' Produced by:
' Portfolio loading process
' PortfolioUpdateService
'
' Consumed by:
' PortfolioUpdateService
' PortfolioValuationService
' RiskEngine
' Reporting
'
' Business invariants:
' - A Position added to the portfolio cannot be Nothing.
' - A CashPosition added to the portfolio cannot be Nothing.
' - A CashPosition must belong to this portfolio.
' - Only one CashPosition may exist per currency.
'
' Mental Model:
' A Portfolio answers:
' "What securities and cash does this investment
'  portfolio currently hold?"
'==========================================================

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pPortfolioId As String
Private pPositions As Collection
Private pCashPositions As Collection

'----------------------------------------------------------
' Class initialization
'
' Every new Portfolio receives two empty collections:
' - one for security positions;
' - one for cash positions.
'----------------------------------------------------------

Private Sub Class_Initialize()

    Set pPositions = New Collection
    Set pCashPositions = New Collection

End Sub

'----------------------------------------------------------
' Portfolio identifier
'----------------------------------------------------------

Public Property Get PortfolioId() As String
    PortfolioId = pPortfolioId
End Property

Public Property Let PortfolioId(ByVal value As String)

    value = Trim$(value)

    If Len(value) = 0 Then
        Err.Raise vbObjectError + 2002, _
            "Portfolio.PortfolioId", _
            "PortfolioId cannot be empty."
    End If

    pPortfolioId = value

End Property

'----------------------------------------------------------
' Security positions collection
'
' Consumers must treat the returned collection as read-only.
' Positions must be added through AddPosition() so that
' portfolio invariants are preserved.
'----------------------------------------------------------

Public Property Get Positions() As Collection
    Set Positions = pPositions
End Property

'----------------------------------------------------------
' Cash positions collection
'
' Consumers must treat the returned collection as read-only.
' Cash positions must be added through AddCashPosition().
'----------------------------------------------------------

Public Property Get CashPositions() As Collection
    Set CashPositions = pCashPositions
End Property

'----------------------------------------------------------
' Add one security position to the portfolio
'----------------------------------------------------------

Public Sub AddPosition(ByVal pos As Position)

    If pos Is Nothing Then
        Err.Raise vbObjectError + 2001, _
            "Portfolio.AddPosition", _
            "Position cannot be Nothing."
    End If

    pPositions.Add pos

End Sub

'----------------------------------------------------------
' Find a security position by instrument identifier
'
' Returns Nothing when no matching position exists.
'----------------------------------------------------------

Public Function FindPosition( _
    ByVal InstrumentId As String) As Position

    Dim pos As Position

    InstrumentId = Trim$(InstrumentId)

    For Each pos In pPositions

        If pos.InstrumentId = InstrumentId Then
            Set FindPosition = pos
            Exit Function
        End If

    Next pos

    Set FindPosition = Nothing

End Function

'----------------------------------------------------------
' Add one cash position to the portfolio
'
' Business rules:
' - The object cannot be Nothing.
' - It must belong to this portfolio.
' - The currency must not already exist.
'----------------------------------------------------------

Public Sub AddCashPosition(ByVal cashPos As CashPosition)

    If cashPos Is Nothing Then
        Err.Raise vbObjectError + 2003, _
            "Portfolio.AddCashPosition", _
            "CashPosition cannot be Nothing."
    End If

    If cashPos.PortfolioId <> pPortfolioId Then
        Err.Raise vbObjectError + 2004, _
            "Portfolio.AddCashPosition", _
            "CashPosition belongs to another portfolio."
    End If

    Dim existingCashPos As CashPosition

    Set existingCashPos = _
        FindCashPosition(cashPos.currencyCode)

    If Not existingCashPos Is Nothing Then
        Err.Raise vbObjectError + 2005, _
            "Portfolio.AddCashPosition", _
            "A cash position already exists for currency " & _
            cashPos.currencyCode & "."
    End If

    pCashPositions.Add cashPos

End Sub

'----------------------------------------------------------
' Find a cash position by currency code
'
' The input is normalized before comparison:
' " eur " becomes "EUR".
'
' Returns Nothing when no matching cash position exists.
'----------------------------------------------------------

Public Function FindCashPosition( _
    ByVal currencyCode As String) As CashPosition

    Dim cashPos As CashPosition
    Dim normalizedCurrencyCode As String

    normalizedCurrencyCode = UCase$(Trim$(currencyCode))

    For Each cashPos In pCashPositions

        If cashPos.currencyCode = normalizedCurrencyCode Then
            Set FindCashPosition = cashPos
            Exit Function
        End If

    Next cashPos

    Set FindCashPosition = Nothing

End Function

'----------------------------------------------------------
' Number of security positions in the portfolio
'----------------------------------------------------------

Public Property Get PositionCount() As Long
    PositionCount = pPositions.Count
End Property

'----------------------------------------------------------
' Number of cash positions in the portfolio
'----------------------------------------------------------

Public Property Get CashPositionCount() As Long
    CashPositionCount = pCashPositions.Count
End Property

