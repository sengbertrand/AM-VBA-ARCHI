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
' Applies an executed order to the portfolio positions.
'
' Producer:
' Consumes Order and ExecutionReport after execution.
'
' Consumer:
' Updates Portfolio and Position objects.
'
' Mental Model:
' This service answers:
' "How does an executed trade change the portfolio?"
'==========================================================
Public Sub ApplyExecution(ByRef ptf As Portfolio, _
                          ByRef ord As Order, _
                          ByRef Report As ExecutionReport)
                              
    '------------------------------------------------------
    ' Pre-conditions
    '------------------------------------------------------
    If ptf Is Nothing Then
        Err.Raise vbObjectError + 3001, _
        "PortfolioUpdateService.ApplyExecution", _
        "Portfolio cannot be Nothing."
    End If
    
    If ord Is Nothing Then
        Err.Raise vbObjectError + 3002, _
        "PortfolioUpdateService.ApplyExecution", _
        "Order cannot be Nothing."
    End If
    
    If Report Is Nothing Then
        Err.Raise vbObjectError + 3003, _
        "PortfolioUpdateService.ApplyExecution", _
        "ExecutionReport cannot be Nothing."
    End If
    
    '------------------------------------------------------
    ' Business consistency check
    '------------------------------------------------------
    
    If ord.PortfolioId <> ptf.PortfolioId Then
        Err.Raise vbObjectError + 3006, _
                  "PortfolioUpdateService.ApplyExecution", _
                  "Order portfolio does not match target portfolio."
    End If
    
    '------------------------------------------------------
    ' Find existing position in the portfolio
    '------------------------------------------------------
    Dim pos As Position
    Set pos = ptf.FindPosition(ord.instrumentId)
    
    If pos Is Nothing Then
    
        Dim posFactory As PositionFactory
        Set posFactory = New PositionFactory
        
        Set pos = posFactory.CreateFromOrder(ptf, ord)
    
        ptf.AddPosition pos
    
        LogInfo "New position created : " & pos.instrumentId
    
    End If
    
    '------------------------------------------------------
    ' Apply quantity impact
    '------------------------------------------------------
    
    If ord.Side = SIDE_BUY Then
    
        pos.Quantity = pos.Quantity + Report.ExecutedQuantity
        
    ElseIf ord.Side = SIDE_SELL Then
        
        pos.Quantity = pos.Quantity - Report.ExecutedQuantity
    
    Else
        
        Err.Raise vbObjectError + 3005, _
            "PortfolioUpdateService.ApplyExecution", _
            "Unsupported order side : " & ord.Side
    
    End If
    
    '------------------------------------------------------
    ' Update latest market price using execution price
    '------------------------------------------------------

    pos.MarketPrice = Report.ExecutionPrice
    
    LogInfo "Portfolio updated : " & ptf.PortfolioId & _
            " | Instrument = " & pos.instrumentId & _
            " | New Qty = " & pos.Quantity & _
            " | New Price = " & pos.MarketPrice

End Sub


