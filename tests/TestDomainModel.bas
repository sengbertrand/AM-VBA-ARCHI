Attribute VB_Name = "TestDomainModel"
Option Explicit


Public Sub TestInvestmentDecision()

    Dim Decision As InvestmentDecision
    Set Decision = New InvestmentDecision
    
    Decision.DecisionId = "DEC_001"
    Decision.PortfolioId = "EURO_BONDS"
    Decision.instrumentId = "OAT_FR_10Y"
    Decision.TargetWeight = 0.15
    Decision.DecisionDate = Date
    Decision.Rationale = "Increase exposure to French government bonds"
    
    LogInfo "InvestmentDecision created : " & Decision.DecisionId
    LogInfo "Target instrument : " & Decision.instrumentId
    LogInfo "Target weight : " & Decision.TargetWeight
    
End Sub

