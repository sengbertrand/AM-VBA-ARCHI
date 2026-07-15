Attribute VB_Name = "Constants"
Option Explicit

'=========================================================
' Module : Constants
'
' Purpose :
'   Centralizes all application-wide constants.
'
' Responsibilities:
'   - Define worksheet names.
'   - Define configuration keys.
'   - Define business constants.
'   - Eliminate hard-coded strings throughout the project.
'=========================================================


Public Const SIDE_BUY As String = "BUY"
Public Const SIDE_SELL As String = "SELL"

Public Const ORDER_CREATED As String = "CREATED"
Public Const ORDER_APPROVED As String = "APPROVED"
Public Const ORDER_REJECTED As String = "REJECTED"
Public Const ORDER_SENT As String = "SENT"
Public Const ORDER_EXECUTED As String = "EXECUTED"

Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARN As String = "WARN"
Public Const LOG_ERROR As String = "ERROR"

Public Const SHEET_LOGS As String = "Logs"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_POSITIONS As String = "Positions"
Public Const SHEET_TRADES As String = "Trades"
Public Const SHEET_INSTRUMENTS As String = "Instruments"
Public Const SHEET_MARKET_DATA As String = "MarketData"
Public Const SHEET_RISK_REPORT As String = "RiskReport"

'----------------------------------------------------------
' Asset classes
'----------------------------------------------------------

Public Const ASSET_CLASS_BOND As String = "BOND"
Public Const ASSET_CLASS_EQUITY As String = "EQUITY"


'----------------------------------------------------------
' Pre-trade risk decisions
'----------------------------------------------------------

Public Const RISK_APPROVED As String = "APPROVED"
Public Const RISK_REJECTED As String = "REJECTED"

'----------------------------------------------------------
' Configuration keys
'----------------------------------------------------------

Public Const CONFIG_BASE_CURRENCY As String = "BaseCurrency"

Public Const CONFIG_MAX_POSITION_WEIGHT As String = _
    "MaximumPositionWeight"

