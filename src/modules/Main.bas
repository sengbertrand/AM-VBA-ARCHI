Attribute VB_Name = "Main"
Option Explicit

'==========================================================
' Module : Main
'
' Classification:
' Application Entry Point
'
' Responsibility:
' Provides the public entry points used to launch the
' application demonstration.
'
' This module must remain lightweight.
'
' Business logic belongs in domain classes and services.
' Test setup belongs in dedicated test modules.
'==========================================================

Public Sub RunDemo()

    LogInfo "------- DEMO STARTED -------"

    TestPreTradeRiskService

    LogInfo "------- DEMO COMPLETED -------"

End Sub
