Attribute VB_Name = "Other_SpeechHelper"
Function CreateSound(xTextToSpeak As String)
    Application.Speech.Speak xTextToSpeak, True
    CreateSound = xTextToSpeak
End Function

Function SymbolToSpeakableName(zOfferID As String)
    zInstrumentName = Range("SymbolIDForInstrument_" & zOfferID).Value
    
    zInstrumentNameBase = Left(zInstrumentName, 3)
    zInstrumentNameCounter = Right(zInstrumentName, 3)
    
    ' Any weird spelling is to help pronunciation
    Select Case zInstrumentNameBase
        Case "USD"
            zInstrumentSpeakableNameBase = "Dollar"
        Case "EUR"
            zInstrumentSpeakableNameBase = "Euro"
        Case "GBP"
            zInstrumentSpeakableNameBase = "Pound"
        Case "JPY"
            zInstrumentSpeakableNameBase = "Yen"
        Case "CHF"
            zInstrumentSpeakableNameBase = "Swiss"
        Case "AUD"
            zInstrumentSpeakableNameBase = "Ozzie"
        Case "NZD"
            zInstrumentSpeakableNameBase = "Kiwi"
        Case "CAD"
            zInstrumentSpeakableNameBase = "Cad"
        Case "SEK"
            zInstrumentSpeakableNameBase = "SEK"
        Case "NOK"
            zInstrumentSpeakableNameBase = "NOK"
        Case "MXN"
            zInstrumentSpeakableNameBase = "Peso"
        Case "PLN"
            zInstrumentSpeakableNameBase = "Zloty"
        Case "SGD"
            zInstrumentSpeakableNameBase = "Sing"
        Case "ZAR"
            zInstrumentSpeakableNameBase = "Rand"
        Case "CZK"
            zInstrumentSpeakableNameBase = "Koruna"
        Case "TRY"
            zInstrumentSpeakableNameBase = "Lira"
        Case "RUB"
            zInstrumentSpeakableNameBase = "Ruble"
        Case "DKK"
            zInstrumentSpeakableNameCounter = "DKK"
        Case Else
            zInstrumentSpeakableNameBase = zInstrumentName
    End Select
    
    ' Any weird spelling is to help pronunciation
    Select Case zInstrumentNameCounter
        Case "USD"
            zInstrumentSpeakableNameCounter = "Dollar"
        Case "EUR"
            zInstrumentSpeakableNameCounter = "Euro"
        Case "GBP"
            zInstrumentSpeakableNameCounter = "Pound"
        Case "JPY"
            zInstrumentSpeakableNameCounter = "Yen"
        Case "CHF"
            zInstrumentSpeakableNameCounter = "Swiss"
        Case "AUD"
            zInstrumentSpeakableNameCounter = "Ozzie"
        Case "NZD"
            zInstrumentSpeakableNameCounter = "Kiwi"
        Case "CAD"
            zInstrumentSpeakableNameCounter = "Cad"
        Case "SEK"
            zInstrumentSpeakableNameCounter = "SEK"
        Case "NOK"
            zInstrumentSpeakableNameCounter = "NOK"
        Case "MXN"
            zInstrumentSpeakableNameCounter = "Peso"
        Case "PLN"
            zInstrumentSpeakableNameCounter = "Zloty"
        Case "SGD"
            zInstrumentSpeakableNameCounter = "Sing"
        Case "ZAR"
            zInstrumentSpeakableNameCounter = "Rand"
        Case "CZK"
            zInstrumentSpeakableNameCounter = "Koruna"
        Case "TRY"
            zInstrumentSpeakableNameCounter = "Lira"
        Case "RUB"
            zInstrumentSpeakableNameCounter = "Ruble"
        Case "DKK"
            zInstrumentSpeakableNameCounter = "DKK"
        Case Else
            zInstrumentSpeakableNameCounter = zInstrumentName
    End Select
    
    'Error handling for CFDs
    If Not zInstrumentSpeakableNameBase = zInstrumentName Then
        SymbolToSpeakableName = zInstrumentSpeakableNameBase & " " & zInstrumentSpeakableNameCounter & " "
    Else
        SymbolToSpeakableName = zInstrumentName
    End If
End Function

