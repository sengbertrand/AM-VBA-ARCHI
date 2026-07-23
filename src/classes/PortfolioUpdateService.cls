VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PortfolioUpdateService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : PortfolioUpdateService
'
' Classification:
' Business Service
'
' Domain:
' Investment / Portfolio / IBOR
'
' Responsibility:
' Applies one validated ExecutionImpact to the security
' and cash positions of a Portfolio.
'
' Consumes:
' Portfolio
' ExecutionImpact
'
' Uses:
' PositionFactory
' CashPositionFactory
'
' Updates:
' Position
' CashPosition
' Portfolio
'
' Business rules:
' - Portfolio cannot be Nothing.
' - ExecutionImpact cannot be Nothing.
' - The execution impact must belong to the target
'   portfolio.
' - A security position is created when it does not exist.
' - A cash position is created when appropriate and absent.
' - A cash debit cannot make the balance negative.
'
' Mental Model:
' This service answers:
' "How should an already calculated execution impact
'  be applied to the portfolio?"
'
' Architectural principle:
' This service does not calculate economic movements.
' It only coordinates their application.
'==========================================================

'----------------------------------------------------------
' Applies one ExecutionImpact to a Portfolio.
'
' Workflow:
' 1. Validate inputs and portfolio consistency.
' 2. Find or prepare the security position.
' 3. Find or prepare the cash position.
' 4. Validate that the complete update can succeed.
' 5. Add newly created objects to the portfolio.
' 6. Apply the security movement.
' 7. Apply the cash movement.
' 8. Update the latest market price.
'
' The validation phase occurs before portfolio mutation
' to reduce the risk of leaving a partially updated state.
'----------------------------------------------------------
Public Sub ApplyExecutionImpact( _
    ByVal ptf As Portfolio, _
    ByVal execImpact As ExecutionImpact)

    '------------------------------------------------------
    ' Pre-conditions
    '------------------------------------------------------
    If ptf Is Nothing Then
        Err.Raise vbObjectError + 3001, _
            "PortfolioUpdateService.ApplyExecutionImpact", _
            "Portfolio cannot be Nothing."
    End If

    If execImpact Is Nothing Then
        Err.Raise vbObjectError + 3002, _
            "PortfolioUpdateService.ApplyExecutionImpact", _
            "ExecutionImpact cannot be Nothing."
    End If

    '------------------------------------------------------
    ' Business consistency
    '------------------------------------------------------
    If execImpact.PortfolioId <> ptf.PortfolioId Then
        Err.Raise vbObjectError + 3006, _
            "PortfolioUpdateService.ApplyExecutionImpact", _
            "ExecutionImpact portfolio does not match target portfolio."
    End If

    '------------------------------------------------------
    ' Find or prepare the security position.
    '
    ' A newly created Position is not immediately added to
    ' the Portfolio. We first validate the complete update.
    ' This avoids leaving an empty Position behind if the
    ' subsequent cash update cannot succeed.
    '------------------------------------------------------
    Dim pos As Position
    Dim isNewPosition As Boolean

    Set pos = ptf.FindPosition(execImpact.InstrumentId)

    If pos Is Nothing Then

        Dim posFactory As PositionFactory
        Set posFactory = New PositionFactory

        Set pos = posFactory.CreateFromExecutionImpact( _
            ptf:=ptf, _
            execImpact:=execImpact)

        isNewPosition = True

    End If

    '------------------------------------------------------
    ' Find or prepare the cash position.
    '
    ' For a sale, an absent cash position may safely be
    ' created with a zero initial balance and then credited.
    '
    ' For a purchase, an absent cash position means that
    ' no available cash exists for the debit.
    '------------------------------------------------------
    Dim cashPos As CashPosition
    Dim isNewCashPosition As Boolean

    Set cashPos = _
        ptf.FindCashPosition(execImpact.currencyCode)

    If cashPos Is Nothing Then

        If execImpact.CashBalanceChange < 0# Then
            Err.Raise vbObjectError + 3007, _
                "PortfolioUpdateService.ApplyExecutionImpact", _
                "No cash position exists for currency " & _
                execImpact.currencyCode & _
                " and the execution requires a cash debit."
        End If

        Dim cashFactory As CashPositionFactory
        Set cashFactory = New CashPositionFactory

        Set cashPos = cashFactory.Create( _
            ptf:=ptf, _
            currencyCode:=execImpact.currencyCode)

        isNewCashPosition = True

    End If

    '------------------------------------------------------
    ' Validate available cash before changing any object.
    '
    ' CashBalanceChange is signed:
    ' - negative value = cash debit;
    ' - positive value = cash credit.
    '------------------------------------------------------
    If execImpact.CashBalanceChange < 0# Then

        Dim requiredCashAmount As Double
        requiredCashAmount = Abs(execImpact.CashBalanceChange)

        If requiredCashAmount > cashPos.Balance Then
            Err.Raise vbObjectError + 3008, _
                "PortfolioUpdateService.ApplyExecutionImpact", _
                "Insufficient cash balance in currency " & _
                execImpact.currencyCode & _
                ". Required: " & CStr(requiredCashAmount) & _
                " | Available: " & CStr(cashPos.Balance)
        End If

    End If

    '------------------------------------------------------
    ' Add newly prepared objects only after all preliminary
    ' validation has succeeded.
    '------------------------------------------------------
    If isNewPosition Then
        ptf.AddPosition pos

        LogInfo _
            "New position created" & _
            " | Portfolio = " & ptf.PortfolioId & _
            " | Instrument = " & pos.InstrumentId
    End If

    If isNewCashPosition Then
        ptf.AddCashPosition cashPos

        LogInfo _
            "New cash position created" & _
            " | Portfolio = " & ptf.PortfolioId & _
            " | Currency = " & cashPos.currencyCode
    End If

    '------------------------------------------------------
    ' Apply the signed security quantity movement.
    '
    ' The BUY/SELL interpretation has already been handled
    ' by ExecutionImpactCalculator.
    '------------------------------------------------------
    pos.Quantity = _
        pos.Quantity + execImpact.SecurityQuantityChange

    '------------------------------------------------------
    ' Update the latest known market price.
    '------------------------------------------------------
    pos.MarketPrice = execImpact.ExecutionPrice

    '------------------------------------------------------
    ' Apply the signed cash movement.
    '
    ' CashPosition exposes explicit Credit and Debit
    ' operations rather than a public Balance setter.
    '------------------------------------------------------
    If execImpact.CashBalanceChange > 0# Then

        cashPos.Credit execImpact.CashBalanceChange

    ElseIf execImpact.CashBalanceChange < 0# Then

        cashPos.Debit Abs(execImpact.CashBalanceChange)

    End If

    '------------------------------------------------------
    ' Operational trace
    '------------------------------------------------------
    LogInfo _
        "Portfolio updated" & _
        " | Portfolio = " & ptf.PortfolioId & _
        " | Instrument = " & pos.InstrumentId & _
        " | Quantity change = " & _
            execImpact.SecurityQuantityChange & _
        " | New quantity = " & pos.Quantity & _
        " | Execution price = " & pos.MarketPrice & _
        " | Currency = " & cashPos.currencyCode & _
        " | Cash change = " & execImpact.CashBalanceChange & _
        " | New cash balance = " & cashPos.Balance

End Sub

