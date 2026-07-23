VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutionImpactCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : ExecutionImpactCalculator
'
' Classification:
' Business Service
'
' Domain:
' Investment / Execution / Portfolio
'
' Responsibility:
' Calculates the economic impacts of an executed order
' before they are applied to the target portfolio.
'
' Inputs:
' Order
' ExecutionReport
' Instrument
'
' Produces:
' ExecutionImpact
'
' Mental Model:
' "What economic changes should this execution
' produce in the portfolio?"
'==========================================================

'----------------------------------------------------------
' Calculates the complete execution impact.
'
' Workflow:
' 1. Validate required inputs.
' 2. Validate document consistency.
' 3. Validate the data required for calculation.
' 4. Calculate the security quantity change.
' 5. Calculate the absolute cash amount.
' 6. Apply the economic direction to cash.
' 7. Build the immutable ExecutionImpact.
'----------------------------------------------------------

Public Function Calculate( _
    ByVal ord As Order, _
    ByVal report As ExecutionReport, _
    ByVal currentInstr As instrument) As ExecutionImpact

    Dim SecurityQuantityChange As Double
    Dim absoluteCashAmount As Double
    Dim CashBalanceChange As Double

    '------------------------------------------------------
    ' 1. Validate the supplied business documents
    '------------------------------------------------------

    ValidateInputs _
        ord:=ord, _
        report:=report, _
        currentInstr:=currentInstr

    '------------------------------------------------------
    ' 2. Validate that the documents describe
    '    the same business operation
    '------------------------------------------------------

    ValidateConsistency _
        ord:=ord, _
        report:=report, _
        currentInstr:=currentInstr

    '------------------------------------------------------
    ' 3. Validate that the data can be calculated
    '------------------------------------------------------

    ValidateCalculationData _
        ord:=ord, _
        report:=report, _
        currentInstr:=currentInstr

    '------------------------------------------------------
    ' 4. Calculate the signed security movement
    '------------------------------------------------------

    SecurityQuantityChange = _
        CalculateSecurityQuantityChange( _
            executedQuantity:=report.executedQuantity, _
            Side:=ord.Side)

    '------------------------------------------------------
    ' 5. Calculate the absolute economic cash amount
    '------------------------------------------------------

    absoluteCashAmount = _
        CalculateAbsoluteCashAmount( _
            executedQuantity:=report.executedQuantity, _
            ExecutionPrice:=report.ExecutionPrice, _
            assetClass:=currentInstr.assetClass)

    '------------------------------------------------------
    ' 6. Apply the economic direction to cash
    '------------------------------------------------------

    CashBalanceChange = _
        CalculateCashBalanceChange( _
            absoluteCashAmount:=absoluteCashAmount, _
            Side:=ord.Side)

    '------------------------------------------------------
    ' 7. Build and return the immutable result
    '------------------------------------------------------

    Set Calculate = _
        CreateExecutionImpact( _
            ord:=ord, _
            report:=report, _
            currentInstr:=currentInstr, _
            SecurityQuantityChange:=SecurityQuantityChange, _
            CashBalanceChange:=CashBalanceChange)

End Function

'----------------------------------------------------------
' Validates that all required business documents exist.
'----------------------------------------------------------

Private Sub ValidateInputs( _
    ByVal ord As Order, _
    ByVal report As ExecutionReport, _
    ByVal currentInstr As instrument)

    If ord Is Nothing Then
        Err.Raise vbObjectError + 8201, _
            "ExecutionImpactCalculator.ValidateInputs", _
            "Order cannot be Nothing."
    End If

    If report Is Nothing Then
        Err.Raise vbObjectError + 8202, _
            "ExecutionImpactCalculator.ValidateInputs", _
            "ExecutionReport cannot be Nothing."
    End If

    If currentInstr Is Nothing Then
        Err.Raise vbObjectError + 8203, _
            "ExecutionImpactCalculator.ValidateInputs", _
            "Instrument cannot be Nothing."
    End If

End Sub

'----------------------------------------------------------
' Validates that all supplied documents describe
' the same business operation.
'----------------------------------------------------------

Private Sub ValidateConsistency( _
    ByVal ord As Order, _
    ByVal report As ExecutionReport, _
    ByVal currentInstr As instrument)

    '------------------------------------------------------
    ' The execution report must belong to the supplied order
    '------------------------------------------------------

    If report.OrderId <> ord.OrderId Then
        Err.Raise vbObjectError + 8204, _
            "ExecutionImpactCalculator.ValidateConsistency", _
            "ExecutionReport does not belong to the supplied Order." & _
            " Expected OrderId: " & ord.OrderId & _
            " | Actual OrderId: " & report.OrderId
    End If

    '------------------------------------------------------
    ' The reference-data instrument must match
    ' the instrument requested by the order
    '------------------------------------------------------

    If currentInstr.InstrumentId <> ord.InstrumentId Then
        Err.Raise vbObjectError + 8205, _
            "ExecutionImpactCalculator.ValidateConsistency", _
            "Instrument does not match the Order instrument." & _
            " Expected InstrumentId: " & ord.InstrumentId & _
            " | Actual InstrumentId: " & currentInstr.InstrumentId
    End If

End Sub

'----------------------------------------------------------
' Validates that the supplied business data can be used
' to calculate one execution impact.
'----------------------------------------------------------

Private Sub ValidateCalculationData( _
    ByVal ord As Order, _
    ByVal report As ExecutionReport, _
    ByVal currentInstr As instrument)

    '------------------------------------------------------
    ' Validate the order side
    '------------------------------------------------------

    If ord.Side <> SIDE_BUY And ord.Side <> SIDE_SELL Then
        Err.Raise vbObjectError + 8206, _
            "ExecutionImpactCalculator.ValidateCalculationData", _
            "Unsupported order side: " & ord.Side
    End If

    '------------------------------------------------------
    ' Validate the executed quantity
    '------------------------------------------------------

    If report.executedQuantity <= 0 Then
        Err.Raise vbObjectError + 8207, _
            "ExecutionImpactCalculator.ValidateCalculationData", _
            "ExecutedQuantity must be strictly positive."
    End If

    '------------------------------------------------------
    ' Validate the execution price
    '------------------------------------------------------

    If report.ExecutionPrice <= 0 Then
        Err.Raise vbObjectError + 8208, _
            "ExecutionImpactCalculator.ValidateCalculationData", _
            "ExecutionPrice must be strictly positive."
    End If

    '------------------------------------------------------
    ' Validate the instrument asset class
    '------------------------------------------------------

    If currentInstr.assetClass <> ASSET_CLASS_BOND And _
       currentInstr.assetClass <> ASSET_CLASS_EQUITY Then

        Err.Raise vbObjectError + 8209, _
            "ExecutionImpactCalculator.ValidateCalculationData", _
            "Unsupported asset class: " & currentInstr.assetClass

    End If

    '------------------------------------------------------
    ' Validate the settlement currency
    '------------------------------------------------------

    If Len(Trim$(currentInstr.currencyCode)) = 0 Then
        Err.Raise vbObjectError + 8210, _
            "ExecutionImpactCalculator.ValidateCalculationData", _
            "Instrument CurrencyCode cannot be empty."
    End If

End Sub

'----------------------------------------------------------
' Calculates the signed security quantity change.
'
' BUY  -> positive quantity change
' SELL -> negative quantity change
'----------------------------------------------------------

Private Function CalculateSecurityQuantityChange( _
    ByVal executedQuantity As Double, _
    ByVal Side As String) As Double

    If Side = SIDE_BUY Then

        CalculateSecurityQuantityChange = executedQuantity

    ElseIf Side = SIDE_SELL Then

        CalculateSecurityQuantityChange = -executedQuantity

    Else

        Err.Raise vbObjectError + 8211, _
            "ExecutionImpactCalculator.CalculateSecurityQuantityChange", _
            "Unsupported order side: " & Side

    End If

End Function

'----------------------------------------------------------
' Calculates the absolute cash amount according to
' the instrument asset-class pricing convention.
'
' BOND:
' Quantity * Price / 100
'
' EQUITY:
' Quantity * Price
'----------------------------------------------------------

Private Function CalculateAbsoluteCashAmount( _
    ByVal executedQuantity As Double, _
    ByVal ExecutionPrice As Double, _
    ByVal assetClass As String) As Double

    If assetClass = ASSET_CLASS_BOND Then

        CalculateAbsoluteCashAmount = _
            executedQuantity * ExecutionPrice / 100#

    ElseIf assetClass = ASSET_CLASS_EQUITY Then

        CalculateAbsoluteCashAmount = _
            executedQuantity * ExecutionPrice

    Else

        Err.Raise vbObjectError + 8212, _
            "ExecutionImpactCalculator.CalculateAbsoluteCashAmount", _
            "Unsupported asset class: " & assetClass

    End If

End Function

'----------------------------------------------------------
' Calculates the signed cash balance change.
'
' BUY  -> negative cash movement
' SELL -> positive cash movement
'----------------------------------------------------------

Private Function CalculateCashBalanceChange( _
    ByVal absoluteCashAmount As Double, _
    ByVal Side As String) As Double

    If absoluteCashAmount <= 0 Then
        Err.Raise vbObjectError + 8213, _
            "ExecutionImpactCalculator.CalculateCashBalanceChange", _
            "AbsoluteCashAmount must be strictly positive."
    End If

    If Side = SIDE_BUY Then

        CalculateCashBalanceChange = -absoluteCashAmount

    ElseIf Side = SIDE_SELL Then

        CalculateCashBalanceChange = absoluteCashAmount

    Else

        Err.Raise vbObjectError + 8214, _
            "ExecutionImpactCalculator.CalculateCashBalanceChange", _
            "Unsupported order side: " & Side

    End If

End Function

'----------------------------------------------------------
' Creates one immutable ExecutionImpact business object.
'----------------------------------------------------------

Private Function CreateExecutionImpact( _
    ByVal ord As Order, _
    ByVal report As ExecutionReport, _
    ByVal currentInstr As instrument, _
    ByVal SecurityQuantityChange As Double, _
    ByVal CashBalanceChange As Double) As ExecutionImpact

    Dim execImpact As ExecutionImpact
    Set execImpact = New ExecutionImpact

    execImpact.Initialize _
        orderIdValue:=ord.OrderId, _
        portfolioIdValue:=ord.PortfolioId, _
        instrumentIdValue:=ord.InstrumentId, _
        currencyCodeValue:=currentInstr.currencyCode, _
        sideValue:=ord.Side, _
        executionPriceValue:=report.ExecutionPrice, _
        securityQuantityChangeValue:=SecurityQuantityChange, _
        cashBalanceChangeValue:=CashBalanceChange

    Set CreateExecutionImpact = execImpact

End Function

