Attribute VB_Name = "TestValuation"
Option Explicit

Public Sub TestValueEngine()
    
    '-----------------------
    ' Create a Position
    '----------------------
    
    Dim pos As Position
    Set pos = New Position
    
    pos.Quantity = 500000
    pos.MarketPrice = 100
    pos.instrumentId = "BUND_DE_10Y"
    
    '-------------------
    ' Retrieve instrument from Security Master
    '-------------------------------------
    
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository
    
    Dim foundInstrument As Instrument
    Set foundInstrument = instrumentRepo.GetInstrument(pos.instrumentId)
    
    '--------------------------------------
    ' Calculate Market Value
    '----------------------------
    
    Dim valEngine As ValuationEngine
    Set valEngine = New ValuationEngine
    
    Dim marketValue As Double
    
    marketValue = valEngine.CalculateMarketValue(pos, _
                            foundInstrument)

    
    Debug.Print "Instrument    : " & foundInstrument.instrumentId
    Debug.Print "Asset Class   : " & foundInstrument.AssetClass
    Debug.Print "Quantity        : " & pos.Quantity
    Debug.Print "Price              : " & pos.MarketPrice
    Debug.Print "Market Value : "; marketValue


End Sub

Public Sub TestPortfolioValuationService()

    '------------------------------------------------------
    ' Create portfolio
    '------------------------------------------------------
    Dim ptf As Portfolio
    Set ptf = New Portfolio

    ptf.PortfolioId = "EURO_BONDS"

    '------------------------------------------------------
    ' Create first position
    '------------------------------------------------------
    Dim posOAT As Position
    Set posOAT = New Position

    posOAT.PortfolioId = ptf.PortfolioId
    posOAT.instrumentId = "OAT_FR_10Y"
    posOAT.Quantity = 1000000
    posOAT.MarketPrice = 99.5
    posOAT.CurrencyCode = "EUR"

    ptf.AddPosition posOAT

    '------------------------------------------------------
    ' Create second position
    '------------------------------------------------------
    Dim posBUND As Position
    Set posBUND = New Position

    posBUND.PortfolioId = ptf.PortfolioId
    posBUND.instrumentId = "BUND_DE_10Y"
    posBUND.Quantity = 500000
    posBUND.MarketPrice = 100
    posBUND.CurrencyCode = "EUR"

    ptf.AddPosition posBUND

    '------------------------------------------------------
    ' Calculate total portfolio market value
    '------------------------------------------------------
    Dim portfolioValuation As PortfolioValuationService
    Set portfolioValuation = New PortfolioValuationService

    Dim totalMarketValue As Double
    totalMarketValue = _
        portfolioValuation.CalculatePortfolioMarketValue(ptf)

    '------------------------------------------------------
    ' Verify result
    '------------------------------------------------------
    Debug.Print "Portfolio Id       : " & ptf.PortfolioId
    Debug.Print "Position Count     : " & ptf.PositionCount
    Debug.Print "OAT Market Value   : 995000"
    Debug.Print "BUND Market Value  : 500000"
    Debug.Print "Total Market Value : " & totalMarketValue

End Sub


