VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Portfolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==========================================================
' Business Object : Portfolio
'
' Classification:
' Business Object
'
' Domain:
' Investment / Portfolio / IBOR
'
' Responsibility:
' Stores a collection of Position objects belonging to
' the same portfolio.
'
' Ownership:
' The Portfolio owns its Positions.
'
' Produced by:
' Portfolio loading process
' PortfolioUpdateService
'
' Consumed by:
' PortfolioUpdateService
' ValuationEngine
' Reporting
'
' Mental Model:
' A Portfolio answers:
' "What does this investment portfolio currently hold?"
'==========================================================

'----------------------------------------------------------
' Private fields
'----------------------------------------------------------

Private pPortfolioId As String
Private pPositions As Collection

'----------------------------------------------------------
' Class initialization
'----------------------------------------------------------
Private Sub Class_Initialize()
    Set pPositions = New Collection
End Sub

'----------------------------------------------------------
' Portfolio identifier
'----------------------------------------------------------

Public Property Get PortfolioId() As String
    PortfolioId = pPortfolioId
End Property

Public Property Let PortfolioId(ByVal value As String)
    pPortfolioId = value
End Property

'----------------------------------------------------------
' Positions collection:
' Returns the portfolio positions.
'
' Consumers must treat the returned collection as read-only.
' Positions must be added through AddPosition() so that
' portfolio invariants are preserved.
'----------------------------------------------------------
'----------------------------------------------------------
Public Property Get Positions() As Collection
    Set Positions = pPositions
End Property

'----------------------------------------------------------
' Add one position to the portfolio
'----------------------------------------------------------

Public Sub AddPosition(ByVal pos As Position)
    
    If pos Is Nothing Then
        Err.Raise vbObjectError + 2001, _
            "Portfolio.AddPosition", _
            "Position cannot be Nothing."
    End If
    
    pPositions.Add pos
    
End Sub

'----------------------------------------------------------
' Find a position by instrument identifier
'----------------------------------------------------------

Public Function FindPosition(ByVal instrumentId As String) As Position
    Dim pos As Position
    
    For Each pos In pPositions
        
        If pos.instrumentId = instrumentId Then
            Set FindPosition = pos
            Exit Function
        End If
    
    Next pos
    
    Set FindPosition = Nothing
        
End Function

'----------------------------------------------------------
' Number of positions in the portfolio
'----------------------------------------------------------

Public Property Get PositionCount() As Long
    PositionCount = pPositions.Count
End Property

