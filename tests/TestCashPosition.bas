Attribute VB_Name = "TestCashPosition"
Option Explicit

'==========================================================
' Test Module : TestCashPosition
'
' Responsibility:
' Verifies the isolated behaviour of the CashPosition
' business object.
'
' Tested behaviours:
' - valid initialization;
' - debit operation;
' - credit operation;
' - insufficient cash rejection;
' - usage before initialization rejection.
'==========================================================

Public Sub TestCashPositionBehaviour()

    On Error GoTo TestFailed

    LogInfo "------- TEST CASH POSITION STARTED -------"

    '------------------------------------------------------
    ' 1. Create and initialize one EUR cash position
    '------------------------------------------------------

    Dim cashPos As CashPosition
    Set cashPos = New CashPosition

    cashPos.Initialize _
        portfolioIdValue:="EURO_BONDS", _
        currencyCodeValue:="eur", _
        initialBalance:=500000#

    '------------------------------------------------------
    ' 2. Verify the initialized state
    '------------------------------------------------------

    AssertStringEquals _
        expectedValue:="EURO_BONDS", _
        actualValue:=cashPos.PortfolioId, _
        assertionName:="PortfolioId must be initialized."

    AssertStringEquals _
        expectedValue:="EUR", _
        actualValue:=cashPos.CurrencyCode, _
        assertionName:="CurrencyCode must be normalized."

    AssertDoubleEquals _
        expectedValue:=500000#, _
        actualValue:=cashPos.Balance, _
        assertionName:="Initial balance must equal 500000."

    '------------------------------------------------------
    ' 3. Debit 100000
    '
    ' Expected balance:
    ' 500000 - 100000 = 400000
    '------------------------------------------------------

    cashPos.Debit 100000#

    AssertDoubleEquals _
        expectedValue:=400000#, _
        actualValue:=cashPos.Balance, _
        assertionName:="Balance after debit must equal 400000."

    '------------------------------------------------------
    ' 4. Credit 20000
    '
    ' Expected balance:
    ' 400000 + 20000 = 420000
    '------------------------------------------------------

    cashPos.Credit 20000#

    AssertDoubleEquals _
        expectedValue:=420000#, _
        actualValue:=cashPos.Balance, _
        assertionName:="Balance after credit must equal 420000."

    '------------------------------------------------------
    ' 5. Verify that an excessive debit is rejected
    '------------------------------------------------------

    TestInsufficientCashRejection cashPos

    '------------------------------------------------------
    ' 6. Verify that an uninitialized object cannot be used
    '------------------------------------------------------

    TestUninitializedCashPositionRejection

    LogInfo "------- TEST CASH POSITION PASSED -------"

    Exit Sub

TestFailed:

    LogError _
        "TestCashPositionBehaviour failed. " & _
        "Error " & Err.Number & ": " & Err.Description

    Err.Raise Err.Number, _
        "TestCashPositionBehaviour", _
        Err.Description

End Sub

'----------------------------------------------------------
' Verify insufficient-cash protection
'----------------------------------------------------------

Private Sub TestInsufficientCashRejection( _
    ByVal cashPos As CashPosition)

    On Error GoTo ExpectedError

    cashPos.Debit 600000#

    Err.Raise vbObjectError + 7101, _
        "TestCashPosition.TestInsufficientCashRejection", _
        "Expected insufficient-cash error was not raised."

ExpectedError:

    If Err.Number = vbObjectError + 7005 Then

        Err.Clear

        AssertDoubleEquals _
            expectedValue:=420000#, _
            actualValue:=cashPos.Balance, _
            assertionName:= _
                "Rejected debit must not modify the balance."

        Exit Sub

    End If

    Err.Raise Err.Number, _
        "TestCashPosition.TestInsufficientCashRejection", _
        Err.Description

End Sub

'----------------------------------------------------------
' Verify initialization protection
'----------------------------------------------------------

Private Sub TestUninitializedCashPositionRejection()

    Dim uninitializedCashPos As CashPosition
    Set uninitializedCashPos = New CashPosition

    On Error GoTo ExpectedError

    Debug.Print uninitializedCashPos.Balance

    Err.Raise vbObjectError + 7102, _
        "TestCashPosition.TestUninitializedCashPositionRejection", _
        "Expected initialization error was not raised."

ExpectedError:

    If Err.Number = vbObjectError + 7006 Then
        Err.Clear
        Exit Sub
    End If

    Err.Raise Err.Number, _
        "TestCashPosition.TestUninitializedCashPositionRejection", _
        Err.Description

End Sub

'----------------------------------------------------------
' String assertion
'----------------------------------------------------------

Private Sub AssertStringEquals( _
    ByVal expectedValue As String, _
    ByVal actualValue As String, _
    ByVal assertionName As String)

    If actualValue <> expectedValue Then
        Err.Raise vbObjectError + 7103, _
            "TestCashPosition.AssertStringEquals", _
            assertionName & _
            " Expected: " & expectedValue & _
            " | Actual: " & actualValue
    End If

End Sub

'----------------------------------------------------------
' Numeric assertion
'----------------------------------------------------------

Private Sub AssertDoubleEquals( _
    ByVal expectedValue As Double, _
    ByVal actualValue As Double, _
    ByVal assertionName As String)

    Const TOLERANCE As Double = 1E-06

    If Abs(actualValue - expectedValue) > TOLERANCE Then
        Err.Raise vbObjectError + 7104, _
            "TestCashPosition.AssertDoubleEquals", _
            assertionName & _
            " Expected: " & CStr(expectedValue) & _
            " | Actual: " & CStr(actualValue)
    End If

End Sub

