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

    Dim Report As ExecutionReport
    Set Report = New ExecutionReport
    
    'Simulated broker execution report.
    Report.OrderId = ord.OrderId
    Report.ExecutedQuantity = ord.Quantity
    Report.ExecutionPrice = 100
    Report.ExecutionTime = Now
    Report.Status = ORDER_EXECUTED
    
    LogInfo "Execution report received : " & Report.OrderId & _
        " | Status = " & Report.Status & _
        " | Executed Qty = " & Report.ExecutedQuantity & _
        " | Price = " & Report.ExecutionPrice
    
    ' Update the OMS order status using the broker execution report.
    
    ord.Status = Report.Status
    LogInfo "Order status updated in OMS : " & ord.OrderId & _
        " | Status = " & ord.Status

    Set ExecuteOrder = Report

End Function

