Attribute VB_Name = "TestExecutionImpact"
Option Explicit

'==========
' Test Module: TestExecutionImpact
'
' Responsibility:
' Verifies the behaviour and invariants of the immutable
' ExecutionImpact business object.
'-----------

Public Sub TestExecutionImpactInitialization()
    
    On Error GoTo TestFailed
    
    LogInfo "----- TEST EXECUTION IMPACT STARTED -------"
    
    '------------
    ' 1. Create the technical object
    '---------------
    
    Dim execImpact As ExecutionImpact
    Set execImpact = New ExecutionImpact
    
    '-------
    ' 2. Initialize one valid buy execution impact
    '
    ' Bond cash impact:
    ' 1 000 000 x 99,5 / 100 =995 000 EUR
    '-----------

    execImpact.Initialize _
        orderIdValue:="ORD_001", _
        portfolioIdValue:="EURO_BONDS", _
        instrumentIdValue:="OAT_FR_10Y", _
        currencyCodeValue:="eur", _
        sideValue:="buy", _
        executionPriceValue:=99.5, _
        securityQuantityChangeValue:=1000000#, _
        cashBalanceChangeValue:=-995000#
    
    '-------
    ' 3. Verify traceability
    '--------------
    
    AssertStringEquals _
        expectedValue:="ORD_001", _
        actualValue:=execImpact.OrderId, _
        assertionName:="OrderId must be preserved."
        
    '---------
    ' 4. Verify targets
    '------------
    
    AssertStringEquals _
        expectedValue:="EURO_BONDS", _
        actualValue:=execImpact.PortfolioId, _
        assertionName:="PortfolioId must be preserved."
    
    AssertStringEquals _
        expectedValue:="OAT_FR_10Y", _
        actualValue:=execImpact.InstrumentId, _
        assertionName:="InstrumentId must be preserved."
    
    AssertStringEquals _
        expectedValue:="EUR", _
        actualValue:=execImpact.currencyCode, _
        assertionName:="CurrencyCode must be normalized."
    
    '----------
    ' 5. Verify execution context.
    '---------
    
    AssertStringEquals _
        expectedValue:=SIDE_BUY, _
        actualValue:=execImpact.Side, _
        assertionName:="Side must be normalized."
    
    AssertDoubleEquals _
        expectedValue:=99.5, _
        actualValue:=execImpact.ExecutionPrice, _
        assertionName:="ExecutionPrice must be preserved."
        
    '------------
    ' 6. Verify calculated impacts
    '---------------
    
    AssertDoubleEquals _
        expectedValue:=1000000#, _
        actualValue:=execImpact.SecurityQuantityChange, _
        assertionName:="BUY security quantity change must be positive."
    
    AssertDoubleEquals _
        expectedValue:=-995000#, _
        actualValue:=execImpact.CashBalanceChange, _
        assertionName:="BUY cash balance change must be negative."
        
    LogInfo "----- TEST EXECUTION IMPACT PASSED ------"
    
    Exit Sub
        
TestFailed:
    
    Dim errorNumber As Long
    Dim errorDescription As String
    
    errorNumber = Err.Number
    errorDescription = Err.Description
    
    LogError _
        "TestExecutionImpactInitialization failed. " & _
        "Error :" & errorNumber & ": " & errorDescription

    Err.Raise errorNumber, _
        "TestExecutionImpact.TestExecutionImpactInitialization", _
        errorDescription
    
    
End Sub

'---------------
' Helper : AssertStringEquals
'
' Verify an expected String value
'-----------

Private Sub AssertStringEquals( _
            ByVal expectedValue As String, _
            ByVal actualValue As String, _
            ByVal assertionName As String)
    
    If expectedValue <> actualValue Then
        Err.Raise vbObjectError + 8101, _
            "TestExecutionImpact.AssertStringEquals", _
            assertionName & _
            " Expected: " & expectedValue & _
            " | Actual: " & actualValue
    End If
    
End Sub

'------------
' Helper : AssertDoubleEquals
'
' Verify an expected Double value
'-----------------------------

Private Sub AssertDoubleEquals( _
            ByVal expectedValue As Double, _
            ByVal actualValue As Double, _
            ByVal assertionName As String)
        
    Const TOLERANCE As Double = 1E-06
    
    If Abs(expectedValue - actualValue) > TOLERANCE Then
        Err.Raise vbObjectError + 8102, _
            "TestExecutionImpact.AssertDoubleEquals", _
            assertionName & _
            " Expected: " & CStr(expectedValue) & _
            " | Actual: " & CStr(actualValue)
    End If
    
End Sub
