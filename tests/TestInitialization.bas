Attribute VB_Name = "TestInitialization"
Option Explicit


Public Sub TestInitializeWorkbook()

    InitializeWorkbook
    LogInfo "WorkBook initialized successfully"
    
End Sub

Public Sub TestSharedFramework()

    LogInfo "Shared Framework started"
    LogWarn "This is a warning example"
    LogError "This is an error example"
    LogInfo "Shared Framework completed"
    
End Sub
