Attribute VB_Name = "TestRepositories"
Option Explicit

Public Sub TestInstrumentRepository()

    '------------------------------------------------------
    ' Create repository
    '------------------------------------------------------
    Dim instrumentRepo As InstrumentRepository
    Set instrumentRepo = New InstrumentRepository

    '------------------------------------------------------
    ' Retrieve instrument from Excel Security Master
    '------------------------------------------------------
    Dim foundInstrument As Instrument
    Set foundInstrument = _
        instrumentRepo.GetInstrument("BUND_DE_10Y")

    '------------------------------------------------------
    ' Verify retrieved reference data
    '------------------------------------------------------
    Debug.Print "Instrument Id   : " & foundInstrument.instrumentId
    Debug.Print "Instrument Name : " & foundInstrument.InstrumentName
    Debug.Print "Asset Class     : " & foundInstrument.AssetClass
    Debug.Print "Currency Code   : " & foundInstrument.CurrencyCode

End Sub

Public Sub TestMarketDataRepository()

    '------------------------
    ' Create repository
    '----------------------------
    
    Dim marketDataRepo As MarketDataRepository
    Set marketDataRepo = New MarketDataRepository
    
    '--------------------
    ' Retrieve current market prices
    '---------------------------
    
    Dim oatPrice As Double
    oatPrice = marketDataRepo.GetMarketPrice("OAT_FR_10Y")
    
    Dim bundPrice As Double
    bundPrice = marketDataRepo.GetMarketPrice("BUND_DE_10Y")
    
    '---------------------------------
    ' Verify results
    '--------------------
    
    Debug.Print "OAT Market Price : " & oatPrice
    Debug.Print "BUND Market Price : " & bundPrice
    
End Sub

Public Sub TestConfigRepository()
    
    '---------------
    ' Create repository
    '-------------------------
    
    Dim configRepo As ConfigRepository
    Set configRepo = New ConfigRepository
    
    '----------
    ' Retrieve configuration values
    '---------
    
    Dim baseCurrency As String
    baseCurrency = configRepo.GetStringValue(CONFIG_BASE_CURRENCY)
    
    Dim maximumPositionWeight As Double
    maximumPositionWeight = _
        configRepo.GetDoubleValue(CONFIG_MAX_POSITION_WEIGHT)
    
    '------------
    ' verify results
    '------
    
    Debug.Print "Base Currency                      : " & baseCurrency
    Debug.Print "Maximum Position Weight    : " & _
                Format(maximumPositionWeight, "0.00%")
    
    
End Sub

