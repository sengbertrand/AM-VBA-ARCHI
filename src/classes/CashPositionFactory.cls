VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashPositionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Factory : CashPositionFactory
'
' Classification:
' Factory
'
' Domain:
' Investment / Portfolio / Cash / IBOR
'
' Responsibility:
' Creates valid CashPosition objects from portfolio
' and currency context.
'
' Used by:
' PortfolioUpdateService
' Portfolio loading process
'
' Produces:
' CashPosition
'
' Business rules:
' - Portfolio cannot be Nothing.
' - CurrencyCode cannot be empty.
' - Initial balance cannot be negative.
' - The created CashPosition belongs to the supplied
'   Portfolio.
'
' Mental Model:
' This factory answers:
' "How should a new CashPosition be initialized?"
'==========================================================

'----------------------------------------------------------
' Creates one initialized CashPosition.
'
' The factory centralizes the construction procedure so
' that its consumers do not need to:
'
' - instantiate CashPosition directly;
' - know how Initialize must be called;
' - manually assign the PortfolioId;
' - normalize the currency code.
'
' Parameters:
' ptf
'   Portfolio that will own the cash position.
'
' currencyCode
'   Currency of the cash balance.
'
' initialBalance
'   Initial cash amount.
'   Defaults to zero when omitted.
'
' Returns:
' A valid and initialized CashPosition.
'----------------------------------------------------------
Public Function Create( _
    ByVal ptf As Portfolio, _
    ByVal currencyCode As String, _
    Optional ByVal initialBalance As Double = 0#) _
    As CashPosition

    '------------------------------------------------------
    ' Pre-condition:
    ' A cash position must belong to an existing portfolio.
    '------------------------------------------------------
    If ptf Is Nothing Then
        Err.Raise vbObjectError + 4101, _
            "CashPositionFactory.Create", _
            "Portfolio cannot be Nothing."
    End If

    '------------------------------------------------------
    ' Normalize the currency code.
    '
    ' Example:
    ' " eur " becomes "EUR".
    '------------------------------------------------------
    currencyCode = UCase$(Trim$(currencyCode))

    If Len(currencyCode) = 0 Then
        Err.Raise vbObjectError + 4102, _
            "CashPositionFactory.Create", _
            "CurrencyCode cannot be empty."
    End If

    '------------------------------------------------------
    ' A newly created cash position cannot start with
    ' a negative balance.
    '------------------------------------------------------
    If initialBalance < 0# Then
        Err.Raise vbObjectError + 4103, _
            "CashPositionFactory.Create", _
            "Initial cash balance cannot be negative."
    End If

    '------------------------------------------------------
    ' Create the CashPosition object.
    '------------------------------------------------------
    Dim cashPos As CashPosition
    Set cashPos = New CashPosition

    '------------------------------------------------------
    ' Build its first valid business state.
    '
    ' The PortfolioId is taken directly from the portfolio
    ' supplied to the factory.
    '------------------------------------------------------
    cashPos.Initialize _
        portfolioIdValue:=ptf.PortfolioId, _
        currencyCodeValue:=currencyCode, _
        initialBalance:=initialBalance

    '------------------------------------------------------
    ' Return the initialized business object.
    '------------------------------------------------------
    Set Create = cashPos

End Function

