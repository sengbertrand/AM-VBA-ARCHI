VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CashPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : CashPosition
'
' Classification:
' Business Object
'
' Domain:
' Investment / Portfolio / Cash / IBOR
'
' Responsibility:
' Represents the cash balance held by one portfolio
' in one specific currency.
'
' Business identity:
' PortfolioId + CurrencyCode
'
' Ownership:
' A Portfolio owns its CashPosition objects.
'
' Produced by:
' Portfolio loading process
' Future cash-position factory
'
' Consumed by:
' Portfolio
' PortfolioUpdateService
' PreTradeRiskService
'
' Invariants:
' - PortfolioId must not be empty.
' - CurrencyCode must not be empty.
' - Initial balance must not be negative.
' - Credit and debit amounts must be strictly positive.
' - A debit must not make the balance negative.
'
' Mental Model:
' A CashPosition answers:
' "How much cash does this portfolio currently hold
'  in this currency?"
'==========================================================

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pPortfolioId As String
Private pCurrencyCode As String
Private pBalance As Double
Private pIsInitialized As Boolean

'----------------------------------------------------------
' Initialize the cash position
'
' VBA does not provide parameterized constructors.
' This method therefore creates the first valid state of
' the business object.
'----------------------------------------------------------

Public Sub Initialize( _
    ByVal portfolioIdValue As String, _
    ByVal currencyCodeValue As String, _
    ByVal initialBalance As Double)

    If pIsInitialized Then
        Err.Raise vbObjectError + 7001, _
            "CashPosition.Initialize", _
            "Cash position has already been initialized."
    End If

    portfolioIdValue = Trim$(portfolioIdValue)

    If Len(portfolioIdValue) = 0 Then
        Err.Raise vbObjectError + 7002, _
            "CashPosition.Initialize", _
            "PortfolioId cannot be empty."
    End If

    currencyCodeValue = UCase$(Trim$(currencyCodeValue))

    If Len(currencyCodeValue) = 0 Then
        Err.Raise vbObjectError + 7003, _
            "CashPosition.Initialize", _
            "CurrencyCode cannot be empty."
    End If

    If initialBalance < 0 Then
        Err.Raise vbObjectError + 7004, _
            "CashPosition.Initialize", _
            "Initial cash balance cannot be negative."
    End If

    pPortfolioId = portfolioIdValue
    pCurrencyCode = currencyCodeValue
    pBalance = initialBalance
    pIsInitialized = True

End Sub

'----------------------------------------------------------
' Portfolio identifier
'----------------------------------------------------------

Public Property Get PortfolioId() As String
    EnsureInitialized
    PortfolioId = pPortfolioId
End Property

'----------------------------------------------------------
' Currency code
'----------------------------------------------------------

Public Property Get CurrencyCode() As String
    EnsureInitialized
    CurrencyCode = pCurrencyCode
End Property

'----------------------------------------------------------
' Current cash balance
'
' The balance can be read publicly, but it cannot be
' replaced directly through a Property Let.
'----------------------------------------------------------

Public Property Get Balance() As Double
    EnsureInitialized
    Balance = pBalance
End Property

'----------------------------------------------------------
' Credit the cash position
'
' A credit increases the available cash balance.
' Examples:
' - proceeds from a security sale;
' - coupon payment;
' - dividend payment;
' - external cash contribution.
'----------------------------------------------------------

Public Sub Credit(ByVal amount As Double)

    EnsureInitialized
    EnsurePositiveAmount amount, "CashPosition.Credit"

    pBalance = pBalance + amount

End Sub

'----------------------------------------------------------
' Debit the cash position
'
' A debit decreases the available cash balance.
' Examples:
' - security purchase;
' - fees;
' - external cash withdrawal.
'----------------------------------------------------------

Public Sub Debit(ByVal amount As Double)

    EnsureInitialized
    EnsurePositiveAmount amount, "CashPosition.Debit"

    If amount > pBalance Then
        Err.Raise vbObjectError + 7005, _
            "CashPosition.Debit", _
            "Insufficient cash balance."
    End If

    pBalance = pBalance - amount

End Sub

'----------------------------------------------------------
' Verify that the object has been initialized
'----------------------------------------------------------

Private Sub EnsureInitialized()

    If Not pIsInitialized Then
        Err.Raise vbObjectError + 7006, _
            "CashPosition", _
            "Cash position must be initialized before use."
    End If

End Sub

'----------------------------------------------------------
' Verify that a cash movement amount is strictly positive
'----------------------------------------------------------

Private Sub EnsurePositiveAmount( _
    ByVal amount As Double, _
    ByVal sourceName As String)

    If amount <= 0 Then
        Err.Raise vbObjectError + 7007, _
            sourceName, _
            "Cash movement amount must be strictly positive."
    End If

End Sub

