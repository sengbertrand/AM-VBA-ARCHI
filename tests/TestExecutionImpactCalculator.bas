Attribute VB_Name = "TestExecutionImpactCalculator"
Option Explicit

'==========================================================
' Test Suite : ExecutionImpactCalculator
'
' Classification:
' Unit / Business Service Tests
'
' Subject under test:
' ExecutionImpactCalculator
'
' Responsibility:
' Verifies that ExecutionImpactCalculator correctly
' transforms an Order, an ExecutionReport and an Instrument
' into one ExecutionImpact.
'
' Testing strategy:
' - Exercise the public Calculate method.
' - Verify observable business results.
' - Do not test private implementation methods directly.
'==========================================================

'----------------------------------------------------------
' Runs all tests related to ExecutionImpactCalculator.
'
' At this stage, the suite contains only the first
' nominal business scenario.
'----------------------------------------------------------
Public Sub RunAllExecutionImpactCalculatorTests()

    LogInfo "--------------------------------------------"
    LogInfo "ExecutionImpactCalculator tests started."
    LogInfo "--------------------------------------------"

    TestCalculateBuyBond

    LogInfo "--------------------------------------------"
    LogInfo "ExecutionImpactCalculator tests completed."
    LogInfo "--------------------------------------------"

End Sub

'----------------------------------------------------------
' Verifies the complete nominal BUY BOND workflow.
'
' Business scenario:
' - BUY 300,000 nominal of a bond.
' - Execution price = 100.
'
' Expected security impact:
' +300,000
'
' Expected cash impact:
' -(300,000 ? 100 / 100)
' = -300,000 EUR
'
' This test crosses the complete public pipeline:
' Inputs
'   -> validation
'   -> security calculation
'   -> cash calculation
'   -> ExecutionImpact construction
'----------------------------------------------------------
Private Sub TestCalculateBuyBond()

    On Error GoTo TestFailed

    Dim ord As Order
    Dim report As ExecutionReport
    Dim instrument As instrument

    Dim impactCalc As ExecutionImpactCalculator
    Dim execImpact As ExecutionImpact

    '------------------------------------------------------
    ' Arrange
    ' Create one completely valid business scenario.
    '------------------------------------------------------
    Set ord = CreateValidOrder()
    Set report = CreateValidExecutionReport()
    Set instrument = CreateValidInstrument()

    Set impactCalc = New ExecutionImpactCalculator

    '------------------------------------------------------
    ' Act
    ' Execute the public contract of the business service.
    '------------------------------------------------------
    Set execImpact = impactCalc.Calculate( _
        ord:=ord, _
        report:=report, _
        instrument:=instrument)

    '------------------------------------------------------
    ' Assert
    ' Verify traceability fields.
    '------------------------------------------------------
    AssertStringEquals _
        expectedValue:="ORD_002", _
        actualValue:=execImpact.orderId, _
        assertionName:="BUY Bond - OrderId"

    AssertStringEquals _
        expectedValue:="EURO_BONDS", _
        actualValue:=execImpact.PortfolioId, _
        assertionName:="BUY Bond - PortfolioId"

    AssertStringEquals _
        expectedValue:="OAT_FR_10Y", _
        actualValue:=execImpact.InstrumentId, _
        assertionName:="BUY Bond - InstrumentId"

    AssertStringEquals _
        expectedValue:="EUR", _
        actualValue:=execImpact.CurrencyCode, _
        assertionName:="BUY Bond - CurrencyCode"

    '------------------------------------------------------
    ' Verify execution context.
    '------------------------------------------------------
    AssertStringEquals _
        expectedValue:=SIDE_BUY, _
        actualValue:=execImpact.side, _
        assertionName:="BUY Bond - Side"

    AssertDoubleEquals _
        expectedValue:=100#, _
        actualValue:=execImpact.executionPrice, _
        assertionName:="BUY Bond - ExecutionPrice"

    '------------------------------------------------------
    ' Verify calculated economic movements.
    '------------------------------------------------------
    AssertDoubleEquals _
        expectedValue:=300000#, _
        actualValue:=execImpact.securityQuantityChange, _
        assertionName:="BUY Bond - SecurityQuantityChange"

    AssertDoubleEquals _
        expectedValue:=-300000#, _
        actualValue:=execImpact.cashBalanceChange, _
        assertionName:="BUY Bond - CashBalanceChange"

    LogInfo "PASS - TestCalculateBuyBond"
    Exit Sub

TestFailed:

    LogError _
        "FAIL - TestCalculateBuyBond" & _
        " | Error " & Err.Number & _
        " | " & Err.Description

End Sub

'==========================================================
' Test Data Builders
'
' These helpers create a valid default business scenario.
'
' Each future test will start from this valid foundation
' and modify only the property relevant to the rule tested.
'==========================================================

'----------------------------------------------------------
' Creates a valid BUY order for the EURO_BONDS portfolio.
'----------------------------------------------------------
Private Function CreateValidOrder() As Order

    Dim ord As Order
    Set ord = New Order

    ord.orderId = "ORD_002"
    ord.PortfolioId = "EURO_BONDS"
    ord.InstrumentId = "OAT_FR_10Y"
    ord.side = SIDE_BUY
    ord.Quantity = 300000#
    ord.Status = ORDER_EXECUTED

    Set CreateValidOrder = ord

End Function

'----------------------------------------------------------
' Creates a valid execution report attached to ORD_002.
'----------------------------------------------------------
Private Function CreateValidExecutionReport() As ExecutionReport

    Dim report As ExecutionReport
    Set report = New ExecutionReport

    report.orderId = "ORD_002"
    report.executedQuantity = 300000#
    report.executionPrice = 100#

    Set CreateValidExecutionReport = report

End Function

'----------------------------------------------------------
' Creates valid reference data for the executed bond.
'----------------------------------------------------------
Private Function CreateValidInstrument() As instrument

    Dim instrument As instrument
    Set instrument = New instrument

    instrument.InstrumentId = "OAT_FR_10Y"
    instrument.assetClass = ASSET_CLASS_BOND
    instrument.CurrencyCode = "EUR"

    Set CreateValidInstrument = instrument

End Function

'==========================================================
' Assertions
'==========================================================

'----------------------------------------------------------
' Verifies that two String values are equal.
'----------------------------------------------------------
Private Sub AssertStringEquals( _
    ByVal expectedValue As String, _
    ByVal actualValue As String, _
    ByVal assertionName As String)

    If actualValue <> expectedValue Then

        Err.Raise vbObjectError + 8301, _
            "TestExecutionImpactCalculator.AssertStringEquals", _
            assertionName & _
            " | Expected: " & expectedValue & _
            " | Actual: " & actualValue

    End If

End Sub

'----------------------------------------------------------
' Verifies that two Double values are equal within
' a small numerical tolerance.
'
' The tolerance protects the test against insignificant
' floating-point representation differences.
'----------------------------------------------------------
Private Sub AssertDoubleEquals( _
    ByVal expectedValue As Double, _
    ByVal actualValue As Double, _
    ByVal assertionName As String)

    Const TOLERANCE As Double = 1E-06

    If Abs(actualValue - expectedValue) > TOLERANCE Then

        Err.Raise vbObjectError + 8302, _
            "TestExecutionImpactCalculator.AssertDoubleEquals", _
            assertionName & _
            " | Expected: " & CStr(expectedValue) & _
            " | Actual: " & CStr(actualValue)

    End If

End Sub

