Attribute VB_Name = "TestPortfolioCash"
Option Explicit

'==========================================================
' Test Module : TestPortfolioCash
'
' Responsibility:
' Verifies that Portfolio correctly owns, stores and
' retrieves CashPosition objects.
'
' Tested behaviours:
' - adding a valid CashPosition;
' - finding a CashPosition by currency;
' - normalizing the searched currency code;
' - rejecting a duplicate currency;
' - rejecting a CashPosition owned by another portfolio.
'==========================================================

Public Sub TestPortfolioCashBehaviour()

    On Error GoTo TestFailed

    LogInfo "------- TEST PORTFOLIO CASH STARTED -------"

    '------------------------------------------------------
    ' 1. Create the portfolio
    '------------------------------------------------------

    Dim ptf As Portfolio
    Set ptf = New Portfolio

    ptf.PortfolioId = "EURO_BONDS"

    '------------------------------------------------------
    ' 2. Create and add one EUR cash position
    '------------------------------------------------------

    Dim eurCashPos As CashPosition
    Set eurCashPos = New CashPosition

    eurCashPos.Initialize _
        portfolioIdValue:="EURO_BONDS", _
        currencyCodeValue:="EUR", _
        initialBalance:=500000#

    ptf.AddCashPosition eurCashPos

    AssertLongEquals _
        expectedValue:=1, _
        actualValue:=ptf.CashPositionCount, _
        assertionName:="Portfolio must contain one cash position."

    '------------------------------------------------------
    ' 3. Find the cash position by currency
    '------------------------------------------------------

    Dim foundCashPos As CashPosition

    Set foundCashPos = ptf.FindCashPosition("EUR")

    AssertObjectIsNotNothing _
        actualObject:=foundCashPos, _
        assertionName:="EUR cash position must be found."

    AssertDoubleEquals _
        expectedValue:=500000#, _
        actualValue:=foundCashPos.Balance, _
        assertionName:="Found EUR cash balance must equal 500000."

    '------------------------------------------------------
    ' 4. Verify search normalization
    '
    ' " eur " must be normalized to "EUR".
    '------------------------------------------------------

    Set foundCashPos = ptf.FindCashPosition(" eur ")

    AssertObjectIsNotNothing _
        actualObject:=foundCashPos, _
        assertionName:= _
            "Cash position search must normalize the currency code."

    '------------------------------------------------------
    ' 5. Verify duplicate-currency rejection
    '------------------------------------------------------

    TestDuplicateCashCurrencyRejection ptf

    '------------------------------------------------------
    ' 6. Verify foreign-portfolio rejection
    '------------------------------------------------------

    TestForeignPortfolioCashRejection ptf

    LogInfo "------- TEST PORTFOLIO CASH PASSED -------"

    Exit Sub

TestFailed:

    Dim errorNumber As Long
    Dim errorDescription As String

    errorNumber = Err.Number
    errorDescription = Err.Description

    LogError _
        "TestPortfolioCashBehaviour failed. " & _
        "Error " & errorNumber & ": " & errorDescription

    Err.Raise errorNumber, _
        "TestPortfolioCash.TestPortfolioCashBehaviour", _
        errorDescription

End Sub

'----------------------------------------------------------
' Verify that the Portfolio refuses two CashPosition objects
' for the same currency.
'----------------------------------------------------------

Private Sub TestDuplicateCashCurrencyRejection( _
    ByVal ptf As Portfolio)

    On Error GoTo ExpectedError

    Dim duplicateCashPos As CashPosition
    Set duplicateCashPos = New CashPosition

    duplicateCashPos.Initialize _
        portfolioIdValue:="EURO_BONDS", _
        currencyCodeValue:="eur", _
        initialBalance:=100000#

    ptf.AddCashPosition duplicateCashPos

    Err.Raise vbObjectError + 7201, _
        "TestPortfolioCash.TestDuplicateCashCurrencyRejection", _
        "Expected duplicate-currency error was not raised."

ExpectedError:

    If Err.Number = vbObjectError + 2005 Then

        Err.Clear

        AssertLongEquals _
            expectedValue:=1, _
            actualValue:=ptf.CashPositionCount, _
            assertionName:= _
                "Rejected duplicate must not change the collection."

        Exit Sub

    End If

    Err.Raise Err.Number, _
        "TestPortfolioCash.TestDuplicateCashCurrencyRejection", _
        Err.Description

End Sub

'----------------------------------------------------------
' Verify that the Portfolio refuses a CashPosition owned
' by another portfolio.
'----------------------------------------------------------

Private Sub TestForeignPortfolioCashRejection( _
    ByVal ptf As Portfolio)

    On Error GoTo ExpectedError

    Dim foreignCashPos As CashPosition
    Set foreignCashPos = New CashPosition

    foreignCashPos.Initialize _
        portfolioIdValue:="GLOBAL_MACRO", _
        currencyCodeValue:="USD", _
        initialBalance:=200000#

    ptf.AddCashPosition foreignCashPos

    Err.Raise vbObjectError + 7202, _
        "TestPortfolioCash.TestForeignPortfolioCashRejection", _
        "Expected foreign-portfolio error was not raised."

ExpectedError:

    If Err.Number = vbObjectError + 2004 Then

        Err.Clear

        AssertLongEquals _
            expectedValue:=1, _
            actualValue:=ptf.CashPositionCount, _
            assertionName:= _
                "Rejected foreign cash position must not be added."

        Exit Sub

    End If

    Err.Raise Err.Number, _
        "TestPortfolioCash.TestForeignPortfolioCashRejection", _
        Err.Description

End Sub

'----------------------------------------------------------
' Verify that an object reference is not Nothing
'----------------------------------------------------------

Private Sub AssertObjectIsNotNothing( _
    ByVal actualObject As Object, _
    ByVal assertionName As String)

    If actualObject Is Nothing Then
        Err.Raise vbObjectError + 7203, _
            "TestPortfolioCash.AssertObjectIsNotNothing", _
            assertionName
    End If

End Sub

'----------------------------------------------------------
' Verify an expected Long value
'----------------------------------------------------------

Private Sub AssertLongEquals( _
    ByVal expectedValue As Long, _
    ByVal actualValue As Long, _
    ByVal assertionName As String)

    If actualValue <> expectedValue Then
        Err.Raise vbObjectError + 7204, _
            "TestPortfolioCash.AssertLongEquals", _
            assertionName & _
            " Expected: " & CStr(expectedValue) & _
            " | Actual: " & CStr(actualValue)
    End If

End Sub

'----------------------------------------------------------
' Verify an expected Double value
'----------------------------------------------------------

Private Sub AssertDoubleEquals( _
    ByVal expectedValue As Double, _
    ByVal actualValue As Double, _
    ByVal assertionName As String)

    Const TOLERANCE As Double = 1E-06

    If Abs(actualValue - expectedValue) > TOLERANCE Then
        Err.Raise vbObjectError + 7205, _
            "TestPortfolioCash.AssertDoubleEquals", _
            assertionName & _
            " Expected: " & CStr(expectedValue) & _
            " | Actual: " & CStr(actualValue)
    End If

End Sub

