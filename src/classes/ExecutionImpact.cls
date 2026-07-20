VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutionImpact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pOrderId As String
Private pPortfolioId As String
Private pInstrumentId As String
Private pCurrencyCode As String

Private pSide As String
Private pExecutionPrice As Double

Private pSecurityQuantityChange As Double
Private pCashBalanceChange As Double

Private pIsInitialized As Boolean

'----------------------------------------------------------
' Initialize the execution impact
'
' Creates the first and only valid state of the object.
' After initialization, the object becomes immutable.
'----------------------------------------------------------

Public Sub Initialize( _
    ByVal orderIdValue As String, _
    ByVal portfolioIdValue As String, _
    ByVal instrumentIdValue As String, _
    ByVal currencyCodeValue As String, _
    ByVal sideValue As String, _
    ByVal executionPriceValue As Double, _
    ByVal securityQuantityChangeValue As Double, _
    ByVal cashBalanceChangeValue As Double)


    '------------------------------------------------------
    ' Lifecycle control
    '------------------------------------------------------

    If pIsInitialized Then
        Err.Raise vbObjectError + 8001, _
            "ExecutionImpact.Initialize", _
            "ExecutionImpact has already been initialized."
    End If

    '------------------------------------------------------
    ' Normalize textual inputs
    '------------------------------------------------------

    orderIdValue = Trim$(orderIdValue)
    portfolioIdValue = Trim$(portfolioIdValue)
    instrumentIdValue = Trim$(instrumentIdValue)
    currencyCodeValue = UCase$(Trim$(currencyCodeValue))
    sideValue = UCase$(Trim$(sideValue))

    '------------------------------------------------------
    ' Target and traceability invariants
    '------------------------------------------------------

    If Len(orderIdValue) = 0 Then
        Err.Raise vbObjectError + 8002, _
            "ExecutionImpact.Initialize", _
            "OrderId cannot be empty."
    End If

    If Len(portfolioIdValue) = 0 Then
        Err.Raise vbObjectError + 8003, _
            "ExecutionImpact.Initialize", _
            "PortfolioId cannot be empty."
    End If

    If Len(instrumentIdValue) = 0 Then
        Err.Raise vbObjectError + 8004, _
            "ExecutionImpact.Initialize", _
            "InstrumentId cannot be empty."
    End If

    If Len(currencyCodeValue) = 0 Then
        Err.Raise vbObjectError + 8005, _
            "ExecutionImpact.Initialize", _
            "CurrencyCode cannot be empty."
    End If

    '------------------------------------------------------
    ' Execution-context invariants
    '------------------------------------------------------

    If sideValue <> SIDE_BUY And sideValue <> SIDE_SELL Then
        Err.Raise vbObjectError + 8006, _
            "ExecutionImpact.Initialize", _
            "Unsupported side: " & sideValue
    End If

    If executionPriceValue <= 0 Then
        Err.Raise vbObjectError + 8007, _
            "ExecutionImpact.Initialize", _
            "ExecutionPrice must be strictly positive."
    End If

    '------------------------------------------------------
    ' Calculated-impact invariants
    '------------------------------------------------------

    If securityQuantityChangeValue = 0 Then
        Err.Raise vbObjectError + 8008, _
            "ExecutionImpact.Initialize", _
            "SecurityQuantityChange cannot be zero."
    End If

    If cashBalanceChangeValue = 0 Then
        Err.Raise vbObjectError + 8009, _
            "ExecutionImpact.Initialize", _
            "CashBalanceChange cannot be zero."
    End If

    If sideValue = SIDE_BUY Then

        If securityQuantityChangeValue < 0 Then
            Err.Raise vbObjectError + 8010, _
                "ExecutionImpact.Initialize", _
                "BUY must produce a positive security quantity change."
        End If

        If cashBalanceChangeValue > 0 Then
            Err.Raise vbObjectError + 8011, _
                "ExecutionImpact.Initialize", _
                "BUY must produce a negative cash balance change."
        End If

    ElseIf sideValue = SIDE_SELL Then

        If securityQuantityChangeValue > 0 Then
            Err.Raise vbObjectError + 8012, _
                "ExecutionImpact.Initialize", _
                "SELL must produce a negative security quantity change."
        End If

        If cashBalanceChangeValue < 0 Then
            Err.Raise vbObjectError + 8013, _
                "ExecutionImpact.Initialize", _
                "SELL must produce a positive cash balance change."
        End If

    End If

    '------------------------------------------------------
    ' Store the validated immutable state
    '------------------------------------------------------

    pOrderId = orderIdValue

    pPortfolioId = portfolioIdValue
    pInstrumentId = instrumentIdValue
    pCurrencyCode = currencyCodeValue

    pSide = sideValue
    pExecutionPrice = executionPriceValue

    pSecurityQuantityChange = securityQuantityChangeValue
    pCashBalanceChange = cashBalanceChangeValue

    pIsInitialized = True

End Sub

Private Sub EnsureInitialized()
    
    If Not pIsInitialized Then
        
        Err.Raise vbObjectError + 8014, _
            "ExecutionImpact.EnsureInitialized", _
            "ExecutionImpact has not been initialized."
    
    End If

End Sub

'----------------------------------------------------------
' Traceability
'----------------------------------------------------------

Public Property Get orderId() As String
    EnsureInitialized
    orderId = pOrderId
End Property

'----------------------------------------------------------
' Targets
'----------------------------------------------------------

Public Property Get PortfolioId() As String
    EnsureInitialized
    PortfolioId = pPortfolioId
End Property

Public Property Get InstrumentId() As String
    EnsureInitialized
    InstrumentId = pInstrumentId
End Property

Public Property Get CurrencyCode() As String
    EnsureInitialized
    CurrencyCode = pCurrencyCode
End Property

'----------------------------------------------------------
' Execution context
'----------------------------------------------------------

Public Property Get Side() As String
    EnsureInitialized
    Side = pSide
End Property

Public Property Get ExecutionPrice() As Double
    EnsureInitialized
    ExecutionPrice = pExecutionPrice
End Property

'----------------------------------------------------------
' Calculated impacts
'----------------------------------------------------------

Public Property Get SecurityQuantityChange() As Double
    EnsureInitialized
    SecurityQuantityChange = pSecurityQuantityChange
End Property

Public Property Get CashBalanceChange() As Double
    EnsureInitialized
    CashBalanceChange = pCashBalanceChange
End Property
