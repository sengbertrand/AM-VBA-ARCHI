VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutionReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'==========================================================
' Business Object : ExecutionReport
'
' Classification:
' Business Object
'
' Domain:
' Investment / OMS / Execution
'
' Responsibility:
' Represents the execution confirmation received from
' a broker, EMS or execution venue.
'
' Producer:
' ExecutionService, broker simulator, EMS interface.
'
' Consumer:
' OMS workflow, Order update process, Portfolio update.
'
' Real-world equivalent:
' Broker execution report / fill confirmation.
'==========================================================

Private pOrderId As String
Private pExecutedQuantity As Double
Private pExecutionPrice As Double
Private pExecutionTime As Date
Private pStatus As String

Public Property Get OrderId() As String
    OrderId = pOrderId
End Property

Public Property Let OrderId(ByVal value As String)
    pOrderId = value
End Property

Public Property Get ExecutedQuantity() As Double
    ExecutedQuantity = pExecutedQuantity
End Property

Public Property Let ExecutedQuantity(ByVal value As Double)
    pExecutedQuantity = value
End Property

Public Property Get ExecutionPrice() As Double
    ExecutionPrice = pExecutionPrice
End Property

Public Property Let ExecutionPrice(ByVal value As Double)
    pExecutionPrice = value
End Property

Public Property Get ExecutionTime() As Date
    ExecutionTime = pExecutionTime
End Property

Public Property Let ExecutionTime(ByVal value As Date)
    pExecutionTime = value
End Property

Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Let Status(ByVal value As String)
    pStatus = value
End Property
