VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutionService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Service : ExecutionService
'
' Classification:
' Business Service
'
' Domain:
' Investment / OMS / Execution
'
' Responsibility:
' Simulates the execution of an order and produces an
' ExecutionReport.
'==========================================================

Public Function ExecuteOrder(ByRef ord As Order) As ExecutionReport

    If ord Is Nothing Then
        Err.Raise vbObjectError + 1001, _
                  "ExecutionService.ExecuteOrder", _
                  "Order cannot be Nothing."
    End If

    'Order is sent to the execution venue.
    ord.Status = ORDER_SENT
    LogInfo "Order sent to execution venue : " & ord.OrderId

    Dim report As ExecutionReport
    Set report = New ExecutionReport
    
    'Simulated broker execution report.
    report.OrderId = ord.OrderId
    report.executedQuantity = ord.Quantity
    report.ExecutionPrice = 100
    report.ExecutionTime = Now
    report.Status = ORDER_EXECUTED
    
    LogInfo "Execution report received : " & report.OrderId & _
        " | Status = " & report.Status & _
        " | Executed Qty = " & report.executedQuantity & _
        " | Price = " & report.ExecutionPrice
    
    ' Update the OMS order status using the broker execution report.
    
    ord.Status = report.Status
    LogInfo "Order status updated in OMS : " & ord.OrderId & _
        " | Status = " & ord.Status

    Set ExecuteOrder = report

End Function

