' ########################################################################################################## '
'                                                                                                            '
' -------------------------------------------------Overview------------------------------------------------- '
' Name:             Account Dashboard                                                                        '
' Dated:            Created: 05.27.13 | Modified 07.13.13                                                    '
' Author:           Jeremy Gulickson                                                                         '
'                                                                                                            '
' Requirements:     Forex Connect + COM integration add on                                                   '
' Download Link:    http://www.fxcodebase.com/wiki/index.php/Download                                        '
' Inspired By:      C:\Program Files (x86)\Candleworks\ForexConnectAPI\samples\com\VBA                       '
'                                                                                                            '
' ----------------------------------------------Version History--------------------------------------------- '
' Version 05.27.13: Created accounts section using table manager                                             '
' Version 06.20.13: Created open trades and closed trades section using table manager                        '
'                   Updated GUI for usability                                                                '
' Version 07.02.13: Added color logic for accounts, open trades and closed trades sections                   '
'                   Updated GUI for usability                                                                '
' Version 07.03.13: Added symbol, base currency, size in usd to open trades and closed trades sections       '
'                   Added sorting to accounts, open trades and closed trades sections                        '
' Version 07.04.13: Added total row to accounts, open trades and closed trades sections                      '
'                   Added color logic and arrow logic to column headings when sorting                        '
'                   Added alternating color background for each row                                          '
'                   Updated leverage calculation to use size in USD for accuracy; previosuly was used margin '
'                   Fixed sorting to work on trades opened and closed after sort is requested                '
' Version 07.05.13: Fixed pip cost calculation for trades where pl = 0                                       '
'                   Updated code for optimization purposes                                                   '
'                   Added extensive code commenting                                                          '
' Version 07.06.13: Added duration open to open trades and closed trades sections                            '
' Version 07.07.13: Added amount and size in usd to account section                                          '
'                   Added counter currency to open and closed trades sections                                '
'                   Added currency section                                                                   '
'                   Added voice for open and closed trades                                                   '
' Version 07.13.13: Updated code for optimization purposes                                                   '
'                                                                                                            '
' ----------------------------------------------Known Concerns---------------------------------------------- '
' - Does not support PIN or sessions selection at login (hardly utilized anyway)                             '
' - Login process is slow on the FXCM network, out of my control                                             '
' - Closing multiple positions rapidly may break formatting but not functionality                            '
' - The upderlying value used to calculate size in usd is only updated at initial login, this is intentional '
'   as typically intraday changes will be irrelveant and therefore is not worth the cpu load                 '
' - Size in usd potentially may not work accurately for CFD products                                         '
' - Size in usd obviously only works for usd denominated accounts                                            '
'                                                                                                            '
' ------------------------------------------------To Do List------------------------------------------------ '
' - Add in system properties values for market status                                                        '
' - Fix issue with partial closes breaking formatting                                                        '
' - Address runtime error produced in currency section                                                       '
'                                                                                                            '
' ########################################################################################################## '


' ########################################################################################################## '
' All variable declarations
' ########################################################################################################## '

Option Explicit

Dim zTransport As Transport
Dim WithEvents zSession As Session
Attribute zSession.VB_VarHelpID = -1
Dim zTableManager As TableManager
Dim zLastSessionStatus As SessionStatusCode

Dim zReadyToStart As Boolean
Dim zTablesLoaded As Boolean
Dim zLoginSuccess As Boolean

Dim i As Integer
Dim k As Integer
Dim zNamedCell As Name
Dim zCell As Range

Dim zDifferenceMinutes As Integer
Dim zDifferenceHours As Integer
Dim zDifferenceDays As Integer
Dim zDifferenceMinutesDisplay As String
Dim zDifferenceHoursDisplay As String
Dim zDifferenceDaysDisplay As String

Dim zCurrencyCollection As Object
Dim zCurrencyRange As Variant

Dim WithEvents zSymbols As OfferTable
Attribute zSymbols.VB_VarHelpID = -1
Dim zSymbolRow As OfferTableRow

Dim WithEvents zAccounts As AccountTable
Attribute zAccounts.VB_VarHelpID = -1
Dim zAccountRow As IAccountTableRow

Dim WithEvents zOpenTrades As TradeTable
Attribute zOpenTrades.VB_VarHelpID = -1
Dim zOpenTradeRow As ITradeTableRow

Dim WithEvents zClosedTrades As ClosedTradeTable
Attribute zClosedTrades.VB_VarHelpID = -1
Dim zClosedTradeRow As IClosedTradeTableRow


' ########################################################################################################## '
' All constant declarations
' ########################################################################################################## '

Const zLengthSymbols As Integer = 5
Const zLengthAccounts As Integer = 12
Const zLengthOpenTrades As Integer = 17
Const zLengthClosedTrades As Integer = 16
Const zLengthCurrency As Integer = 12

' Sets text to speak for open and closed positions
Const zTextToSpeakOpenTrades As String = "Trade Opened"
Const zTextToSpeakClosedTrades As String = "Trade Closed"

' See http://office.microsoft.com/en-us/excel-help/number-format-codes-HP005198679.aspx
Const zNumberFormatText As String = "@"
Const zNumberFormatNumber As String = "0.00"
Const zNumberFormatPL As String = "#,##0.0_);(#,##0.0)"
Const zNumberFormatComma As String = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
Const zNumberFormatUSD As String = "#,##0_);(#,##0)"
Const zNumberFormatTime As String = "mm/dd/yyyy hh:mm"
Const zNumberFormatPercentage As String = "#0.00%_);(#0.00%)"
Const zNumberFormatCurrency As String = "#,##0.00_);(#,##0.00)"
' Saved for potential future use; same as above but with leading '$'
' Const zNumberFormatUSD As String = "$  #,##0_);$ (#,##0)"
' Const zNumberFormatCurrency As String = "$  #,##0.00_);$ (#,##0.00)"


' ########################################################################################################## '
' Code Start
' ########################################################################################################## '

Public Sub Login()
    On Error GoTo Catch
    
    Call SetGlobalVariables
    Call CheckCurrentState
    Call NoErrorMessageText
    
    ' Checks to make sure all the necessary fields is filled out
    ' If any inputs are missing, code will stop and display an error message
    If zLastSessionStatus = SessionStatusCode_Connected Then
        Range("iMessage").Value = "You must log out first."
        Call ErrorMessageText
        Exit Sub
    End If
    If Range("iPassword").Value = vbNullString Then
        Range("iMessage").Value = "Need password."
        Call ErrorMessageText
        Exit Sub
    End If
    If Range("iLogin").Value = vbNullString Then
        Range("iMessage").Value = "Need login."
        Call ErrorMessageText
        Exit Sub
    End If
    If Range("iConnection").Value = vbNullString Then
        Range("iMessage").Value = "Need Connection."
        Call ErrorMessageText
        Exit Sub
    End If
    
    ' Starts connection using Table Manager
    Call zSession.useTableManager(TableManagerMode_Yes, Nothing)
    Call zSession.Login( _
        Range("iLogin").Value, _
        Range("iPassword").Value, _
        "http://fxcorporate.com/Hosts.jsp", _
        Range("iConnection").Value)

    ' Pauses the workflow until login is confirmed
    ' Necessary as log in time is variable and dependent on many factors
    zReadyToStart = False
    Do While Not zReadyToStart
        DoEvents
    Loop
    
    ' Checks to ensure that the tables loaded correctly
    If zLastSessionStatus = SessionStatusCode_Connected And zTablesLoaded Then
        'Application.DisplayFullScreen = True
        Set zTableManager = zSession.getTableManager()
        Set zSymbols = zTableManager.getTable(TableType_Offers)
        Set zAccounts = zTableManager.getTable(TableType_Accounts)
        Set zOpenTrades = zTableManager.getTable(TableType_Trades)
        Set zClosedTrades = zTableManager.getTable(TableType_ClosedTrades)
        Call SpeedUpVBA
        Call ShowSymbols
        Call ShowAccounts
        Call ShowOpenTrades
        Call ShowClosedTrades
        Call ShowCurrencySummary
        Range("iStart").Select

    ' As this uses table manager to retrieve data,
    ' If the tables fail to load the code will stop and display an error message
    ElseIf Not zLoginSuccess Then
        Call ErrorMessageText
        Call ReturnToNormalVBA
        Exit Sub
    ElseIf Not zTablesLoaded Then
        Range("iMessage").Value = "Trading Tables Failed to Load."
        Call ErrorMessageText
        Call ReturnToNormalVBA
        Exit Sub
    End If
    Exit Sub
Catch:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error located in sub Login()"
End Sub

Public Sub Logout()
    On Error Resume Next
    
    Call SetGlobalVariables

    ' Allows the sheet to return to a "clean slate" regardless of the current state
    If zLastSessionStatus <> SessionStatusCode_Connected Then
        Call ErrorMessageText
        With Range("iConnectionStatus")
            .Value = "Disconnected"
            .Interior.Color = zColorNegative
        End With
        Range("iMessage").Value = "You are not logged in."
    Else
        Call NoErrorMessageText
        Call zSession.Logout
    End If
    
    ' Clears section heading number values
    Range("aAccountsData") = "Accounts Data ()"
    Range("tOpenTradesData") = "Open Trades Data ()"
    Range("cClosedTradesData") = "Closed Trades Data ()"
    Range("uCurrencyData") = "Currency Data ()"
    
    ' Clears sorting for each section
    Range("aAccountName").Resize(, zLengthAccounts).Interior.Color = vbBlack
    Range("tAccountName").Resize(, zLengthOpenTrades).Interior.Color = vbBlack
    Range("cAccountName").Resize(, zLengthClosedTrades).Interior.Color = vbBlack
    Range("uAccountName").Resize(, zLengthCurrency).Interior.Color = vbBlack
    For Each zCell In Range("aAccountName").Resize(, zLengthAccounts)
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
    Next zCell
    For Each zCell In Range("tAccountName").Resize(, zLengthOpenTrades)
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
    Next zCell
    For Each zCell In Range("cAccountName").Resize(, zLengthClosedTrades)
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
    Next zCell
    For Each zCell In Range("uAccountName").Resize(, zLengthCurrency)
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
        If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
    Next zCell
    
    ' Clears select named cells
    For Each zNamedCell In Names
        If InStr(1, zNamedCell.Name, "SymbolIDForInstrument_") > 0 Then Range(zNamedCell.Name).Resize(, zLengthSymbols).ClearContents
        If InStr(1, zNamedCell.Name, "SymbolID") > 0 Then zNamedCell.Delete
        If InStr(1, zNamedCell.Name, "Total_") > 0 Then Range(zNamedCell.Name).EntireRow.Delete
        If InStr(1, zNamedCell.Name, "Account_") > 0 Then Range(zNamedCell.Name).EntireRow.Delete
        If InStr(1, zNamedCell.Name, "OpenTicket_") > 0 Then Range(zNamedCell.Name).EntireRow.Delete
        If InStr(1, zNamedCell.Name, "ClosedTicket_") > 0 Then Range(zNamedCell.Name).EntireRow.Delete
        If InStr(1, zNamedCell.Name, "CurrencySummary_") > 0 Then Range(zNamedCell.Name).EntireRow.Delete
        If InStr(1, zNamedCell.RefersTo, "#REF!") > 0 Then zNamedCell.Delete
    Next zNamedCell
    
    Call ReturnToNormalVBA
    Range("iStart").Select
End Sub

Private Sub ShowSymbols()
    ' Offers are not actually utilized anywhere this simply allows instrument name instead of Offer ID
    ' Also used to calculate size in usd
    On Error GoTo Catch
    For i = 0 To zSymbols.Count - 1
        Set zSymbolRow = zSymbols.getRow(i)
        Range("oInstrument").Offset(i + 1, 0).Value = zSymbolRow.Instrument
        Names.Add Name:="SymbolIDForInstrument_" & zSymbolRow.OfferID, RefersTo:=Range("oInstrument").Offset(i + 1)
        Range("oContractCurrency").Offset(i + 1, 0).Value = zSymbolRow.ContractCurrency
        Names.Add Name:="SymbolIDForContractCurrency_" & zSymbolRow.OfferID, RefersTo:=Range("oContractCurrency").Offset(i + 1)
        If zSymbolRow.InstrumentType = "1" Then
            Range("oCounterCurrency").Offset(i + 1, 0).Value = Right(zSymbolRow.Instrument, 3)
            Names.Add Name:="SymbolIDForCounterCurrency_" & zSymbolRow.OfferID, RefersTo:=Range("oCounterCurrency").Offset(i + 1)
        Else
            Range("oCounterCurrency").Offset(i + 1, 0).Value = ""
            Names.Add Name:="SymbolIDForCounterCurrency_" & zSymbolRow.OfferID, RefersTo:=Range("oCounterCurrency").Offset(i + 1)
        End If
        Range("oBid").Offset(i + 1, 0).Value = zSymbolRow.Bid
        Names.Add Name:="SymbolIDForRate_" & zSymbolRow.OfferID, RefersTo:=Range("oBid").Offset(i + 1)
        Range("oInstrumentType").Offset(i + 1, 0).Value = zSymbolRow.InstrumentType
        Names.Add Name:="SymbolIDForInstrumentType_" & zSymbolRow.OfferID, RefersTo:=Range("oInstrumentType").Offset(i + 1)
    Next i
    For i = 0 To zSymbols.Count - 1
        Set zSymbolRow = zSymbols.getRow(i)
        If zSymbolRow.InstrumentType = "1" Then
        ' Calculates and populates "Size in USD" for FX symbols
            With Range("oSizeInUSD").Offset(i + 1)
                If zSymbolRow.ContractCurrency = "AUD" Then
                    .Value = Round(Application.WorksheetFunction.VLookup("AUD/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "CAD" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/CAD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "CHF" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/CHF", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "EUR" Then
                    .Value = Round(Application.WorksheetFunction.VLookup("EUR/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "GBP" Then
                    .Value = Round(Application.WorksheetFunction.VLookup("GBP/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "HKD" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/HKD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "JPY" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/JPY", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "NOK" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/NOK", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "NZD" Then
                    .Value = Round(Application.WorksheetFunction.VLookup("NZD/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "SEK" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/SEK", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "SGD" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/SGD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "TRY" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/TRY", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                ElseIf zSymbolRow.ContractCurrency = "USD" Then
                    .Value = 1000
                ElseIf zSymbolRow.ContractCurrency = "ZAR" Then
                    .Value = Round(1 / Application.WorksheetFunction.VLookup("USD/ZAR", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * 1000, 0)
                End If
            End With
        Else
        ' Calculates and populates "Size in USD" for all non-FX symbols
            With Range("oSizeInUSD").Offset(i + 1)
                If zSymbolRow.ContractCurrency = "AUD" Then
                    .Value = Round(zSymbolRow.Bid * Application.WorksheetFunction.VLookup("AUD/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "CAD" Then
                    .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/CAD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "CHF" Then
                    .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/CHF", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "EUR" Then
                    .Value = Round(zSymbolRow.Bid * Application.WorksheetFunction.VLookup("EUR/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "GBP" Then
                    .Value = Round(zSymbolRow.Bid * Application.WorksheetFunction.VLookup("GBP/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "HKD" Then
                    ' HKG33 does not correctly populate the "ContractMultiplier" field so this adjusts for it
                    If zSymbolRow.Instrument <> "HKG33" Then
                        .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/HKD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                    Else
                        .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/HKD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier * 10, 0)
                    End If
                ElseIf zSymbolRow.ContractCurrency = "JPY" Then
                    ' JPN225 does not correctly populate the "ContractMultiplier" field so this adjusts for it
                    If zSymbolRow.Instrument <> "JPN225" Then
                        .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/JPY", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                    Else
                        .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/JPY", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier * 100, 0)
                    End If
                ElseIf zSymbolRow.ContractCurrency = "NOK" Then
                    .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/NOK", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "NZD" Then
                    .Value = Round(zSymbolRow.Bid * Application.WorksheetFunction.VLookup("NZD/USD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "SEK" Then
                    .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/SEK", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "SGD" Then
                    .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/SGD", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "TRY" Then
                    .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/TRY", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                ElseIf zSymbolRow.ContractCurrency = "USD" Then
                    ' Silver does not correctly populate the "ContractMultiplier" field so this adjusts for it
                    If zSymbolRow.Instrument <> "XAG/USD" Then
                        .Value = zSymbolRow.Bid * zSymbolRow.ContractMultiplier
                    Else
                        .Value = zSymbolRow.Bid * 50
                    End If
                ElseIf zSymbolRow.ContractCurrency = "ZAR" Then
                    .Value = Round(zSymbolRow.Bid * (1 / Application.WorksheetFunction.VLookup("USD/ZAR", Range(Range("oInstrument").Offset(1, 0).Address & ":" & Range("oSizeInUSD").Offset(zSymbols.Count, 0).Address), 4, False)) * zSymbolRow.ContractMultiplier, 0)
                End If
            End With
        End If
        Names.Add Name:="SymbolIDForSizeInUSD_" & zSymbolRow.OfferID, RefersTo:=Range("oSizeInUSD").Offset(i + 1)
    Next i
    Exit Sub
Catch:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error located in sub ShowSymbols()"
End Sub

Private Sub ShowAccounts()
    On Error GoTo Catch
    
    ' Sets section column heading value
    Range("aAccountsData") = "Accounts Data (" & zAccounts.Count & ")"
        
    ' Inserts the number of rows necessary to display all accounts + a total row
    For i = 0 To zAccounts.Count
        Range("aAccountName").Offset(2, 0).EntireRow.Insert
    Next i
    
    ' Iterates through all accounts
    For i = 0 To zAccounts.Count - 1
        Set zAccountRow = zAccounts.getRow(i)
        
        'Applies formatting
        Range("aAccountName").Offset(i + 1, 0).Resize(, zLengthAccounts).Font.Color = vbBlack
        Range("aAccountName").Offset(i + 1, 0).NumberFormat = zNumberFormatText
        Range("aEquity").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("aAmount").Offset(i + 1, 0).NumberFormat = zNumberFormatUSD
        Range("aSizeInUSD").Offset(i + 1, 0).NumberFormat = zNumberFormatUSD
        Range("aLeverage").Offset(i + 1, 0).NumberFormat = zNumberFormatNumber
        Range("aGrossPL").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("aGrossPLp").Offset(i + 1, 0).NumberFormat = zNumberFormatPercentage
        Range("aDayPL").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("aDayPLp").Offset(i + 1, 0).NumberFormat = zNumberFormatPercentage
        Range("aUsedMargin").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("aBaseUnitSize").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        
        ' Names cell for future reference
        Names.Add Name:="Account_" & zAccountRow.AccountName, RefersTo:=Range("aAccountName").Offset(i + 1)
        
         ' Populates values
        Range("aAccountName").Offset(i + 1, 0).Value = zAccountRow.AccountName
        Range("aEquity").Offset(i + 1, 0).Value = zAccountRow.Equity
        Range("aGrossPL").Offset(i + 1, 0).Value = zAccountRow.GrossPL
        Range("aGrossPLp").Offset(i + 1, 0).Value = (zAccountRow.GrossPL / zAccountRow.Equity)
        Range("aDayPL").Offset(i + 1, 0).Value = zAccountRow.DayPL
        Range("aDayPLp").Offset(i + 1, 0).Value = (zAccountRow.DayPL / zAccountRow.Equity)
        ' Below function was removed from ForexConnect in 2016, need to sum individual trade rows
        ' Range("aUsedMargin").Offset(i + 1, 0).Value = zAccountRow.UsedMargin
        Range("aUsedMargin").Offset(i + 1, 0).Value = vbNullString
        Range("aBaseUnitSize").Offset(i + 1, 0).Value = zAccountRow.BaseUnitSize
                
        ' Colors gross p/l & percentage text
        If zAccountRow.GrossPL < 0 Then
            Range("aGrossPL").Offset(i + 1, 0).Font.Color = zColorNegative
            Range("aGrossPLp").Offset(i + 1, 0).Font.Color = zColorNegative
        ElseIf zAccountRow.DayPL = 0 Then
            Range("aGrossPL").Offset(i + 1, 0).Font.Color = zColorNeutral
            Range("aGrossPLp").Offset(i + 1, 0).Font.Color = zColorNeutral
        Else
            Range("aGrossPL").Offset(i + 1, 0).Font.Color = zColorPositive
            Range("aGrossPLp").Offset(i + 1, 0).Font.Color = zColorPositive
        End If
        
        ' Colors day p/l & percentage text
        If zAccountRow.DayPL < 0 Then
            Range("aDayPL").Offset(i + 1, 0).Font.Color = zColorNegative
            Range("aDayPLp").Offset(i + 1, 0).Font.Color = zColorNegative
        ElseIf zAccountRow.DayPL = 0 Then
            Range("aDayPL").Offset(i + 1, 0).Font.Color = zColorNeutral
            Range("aDayPLp").Offset(i + 1, 0).Font.Color = zColorNeutral
        Else
            Range("aDayPL").Offset(i + 1, 0).Font.Color = zColorPositive
            Range("aDayPLp").Offset(i + 1, 0).Font.Color = zColorPositive
        End If
    Next i
    
    ' Names total row cells for future reference and applies formatting
    Names.Add Name:="Total_aAccountName", RefersTo:=Range("aAccountName").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aEquity", RefersTo:=Range("aEquity").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aAmount", RefersTo:=Range("aAmount").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aSizeInUSD", RefersTo:=Range("aSizeInUSD").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aLeverage", RefersTo:=Range("aLeverage").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aGrossPL", RefersTo:=Range("aGrossPL").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aGrossPLp", RefersTo:=Range("aGrossPLp").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aDayPL", RefersTo:=Range("aDayPL").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aDayPLp", RefersTo:=Range("aDayPLp").Offset(zAccounts.Count + 1, 0)
    Names.Add Name:="Total_aUsedMargin", RefersTo:=Range("aUsedMargin").Offset(zAccounts.Count + 1, 0)
    Range("Total_aEquity").NumberFormat = zNumberFormatCurrency
    Range("Total_aAmount").NumberFormat = zNumberFormatUSD
    Range("Total_aSizeInUSD").NumberFormat = zNumberFormatUSD
    Range("Total_aLeverage").NumberFormat = zNumberFormatNumber
    Range("Total_aGrossPL").NumberFormat = zNumberFormatCurrency
    Range("Total_aGrossPLp").NumberFormat = zNumberFormatPercentage
    Range("Total_aDayPL").NumberFormat = zNumberFormatCurrency
    Range("Total_aDayPLp").NumberFormat = zNumberFormatPercentage
    Range("Total_aUsedMargin").NumberFormat = zNumberFormatCurrency
    With Range("Total_aAccountName")
        .Resize(, zLengthAccounts).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Resize(, zLengthAccounts).Borders(xlEdgeTop).Color = zColorTotalLine
        .Resize(, zLengthAccounts).Interior.Color = zColorTotalRow
    End With
    
    Call ColorAccounts
    Call RefreshTotalAccounts
    Exit Sub
Catch:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error located in sub ShowAccounts()"
End Sub

Private Sub zAccounts_Changed(ByVal xRowID As String, ByVal xAccountRowMessage As fxcore2_com.IRow)
    On Error Resume Next
    Set zAccountRow = xAccountRowMessage
    
    ' Determines where to place the updated values
    i = Range("Account_" & zAccountRow.AccountName).Row - Range("aAccountName").Row
    
    ' Populates values
    Range("aEquity").Offset(i, 0).Value = zAccountRow.Equity
    Range("aGrossPL").Offset(i, 0).Value = zAccountRow.GrossPL
    Range("aGrossPLp").Offset(i, 0).Value = (zAccountRow.GrossPL / zAccountRow.Equity)
    Range("aDayPL").Offset(i, 0).Value = zAccountRow.DayPL
    Range("aDayPLp").Offset(i, 0).Value = (zAccountRow.DayPL / zAccountRow.Equity)
    ' Below function was removed from ForexConnect in 2016, need to sum individual trade rows
    ' Range("aUsedMargin").Offset(i, 0).Value = zAccountRow.UsedMargin
    Range("aUsedMargin").Offset(i, 0).Value = vbNullString
    
    'Performs amount calculation using the open trades section
    Range("aAmount").Offset(i, 0).Value = Application.WorksheetFunction.SumIf( _
        Range(Range("tAccountName").Offset(1, 0).Address & ":" & Range("tAccountName").Offset(zOpenTrades.Count, 0).Address), _
        zAccountRow.AccountName, _
        Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address))
    
    'Performs size in usd calculation using the open trades section
    Range("aSizeInUSD").Offset(i, 0).Value = Application.WorksheetFunction.SumIf( _
        Range(Range("tAccountName").Offset(1, 0).Address & ":" & Range("tAccountName").Offset(zOpenTrades.Count, 0).Address), _
        zAccountRow.AccountName, _
        Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address))

    ' Performs an acccurate leverage calculation
    Range("aLeverage").Offset(i, 0).Value = Range("aSizeInUSD").Offset(i, 0).Value / zAccountRow.Equity
    
    ' Colors gross p/l & percentage text
    If zAccountRow.GrossPL < 0 Then
        Range("aGrossPL").Offset(i, 0).Font.Color = zColorNegative
        Range("aGrossPLp").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zAccountRow.DayPL = 0 Then
        Range("aGrossPL").Offset(i, 0).Font.Color = zColorNeutral
        Range("aGrossPLp").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("aGrossPL").Offset(i, 0).Font.Color = zColorPositive
        Range("aGrossPLp").Offset(i, 0).Font.Color = zColorPositive
    End If
    
    ' Colors day p/l & percentage text
    If zAccountRow.DayPL < 0 Then
        Range("aDayPL").Offset(i, 0).Font.Color = zColorNegative
        Range("aDayPLp").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zAccountRow.DayPL = 0 Then
        Range("aDayPL").Offset(i, 0).Font.Color = zColorNeutral
        Range("aDayPLp").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("aDayPL").Offset(i, 0).Font.Color = zColorPositive
        Range("aDayPLp").Offset(i, 0).Font.Color = zColorPositive
    End If
    
    Call RefreshTotalAccounts
End Sub

Private Sub RefreshTotalAccounts()
    On Error Resume Next
    If zAccounts.Count > 0 Then
        
        ' Calculates total equity
        Range("Total_aEquity").Value = Application.WorksheetFunction.Sum( _
            Range(Range("aEquity").Offset(1, 0).Address & ":" & Range("aEquity").Offset(zAccounts.Count, 0).Address))
        
        ' Calculates total amount
        With Range("Total_aAmount")
            If Not Range("Total_tAmount").Value = vbNullString Then
                .Value = Range("Total_tAmount").Value
            Else
                .Value = 0
            End If
        End With
        
        ' Calculates total size in usd
        With Range("Total_aSizeInUSD")
            If Not Range("Total_tSizeInUSD").Value = vbNullString Then
                .Value = Range("Total_tSizeInUSD").Value
            Else
                .Value = 0
            End If
        End With
        
        ' Calculates total leverage
        Range("Total_aLeverage").Value = Range("Total_tSizeInUSD").Value / Range("Total_aEquity").Value
        
        ' Calculates and colors total gross p/l
        With Range("Total_aGrossPL")
            .Value = Application.WorksheetFunction.Sum(Range(Range("aGrossPL").Offset(1, 0).Address & ":" _
                & Range("aGrossPL").Offset(zAccounts.Count, 0).Address))
            If .Value < 0 Then
                .Font.Color = zColorNegative
            ElseIf .Value = 0 Then
                .Font.Color = zColorNeutral
            Else
                .Font.Color = zColorPositive
            End If
        End With
        
        ' Calculates and colors total gross p/l %
        With Range("Total_aGrossPLp")
            .Value = Range("Total_aGrossPL").Value / Range("Total_aEquity").Value
            If .Value < 0 Then
                .Font.Color = zColorNegative
            ElseIf .Value = 0 Then
                .Font.Color = zColorNeutral
            Else
                .Font.Color = zColorPositive
            End If
        End With
        
        ' Calculates and colors total day p/l
        With Range("Total_aDayPL")
            .Value = Application.WorksheetFunction.Sum(Range(Range("aDayPL").Offset(1, 0).Address & ":" _
                & Range("aDayPL").Offset(zAccounts.Count, 0).Address))
            If .Value < 0 Then
                .Font.Color = zColorNegative
            ElseIf .Value = 0 Then
                .Font.Color = zColorNeutral
            Else
                .Font.Color = zColorPositive
            End If
        End With
        
        ' Calculates and colors total day p/l %
       With Range("Total_aDayPLp")
            .Value = Range("Total_aDayPL").Value / Range("Total_aEquity").Value
            If .Value < 0 Then
                .Font.Color = zColorNegative
            ElseIf .Value = 0 Then
                .Font.Color = zColorNeutral
            Else
                .Font.Color = zColorPositive
            End If
        End With
        
        ' Calculates total user margin
        Range("Total_aUsedMargin").Value = Application.WorksheetFunction.Sum( _
            Range(Range("aUsedMargin").Offset(1, 0).Address & ":" & Range("aUsedMargin").Offset(zAccounts.Count, 0).Address))
    Else
        ' If there are no accounts, this clears any existing formatting and sets values to 0 in the total row
        Range("Total_aEquity").Value = vbNullString
        Range("Total_aGrossPL").Value = vbNullString
        Range("Total_aGrossPLp").Value = vbNullString
        Range("Total_aDayPL").Value = vbNullString
        Range("Total_aDayPLp").Value = vbNullString
        Range("Total_aUsedMargin").Value = vbNullString
        Range("Total_aAccountName").Resize(, zLengthAccounts).Interior.Color = xlNone
    End If
End Sub

' Necessary as amount and size in usd are calculated using open trades section which is loaded after accounts section
Private Sub AccountsHelper()

    ' Iterates through all accounts
    For i = 1 To zAccounts.Count
        'Performs an amount calculation using the open trades section
        Range("aAmount").Offset(i, 0).Value = Application.WorksheetFunction.SumIf( _
            Range(Range("tAccountName").Offset(1, 0).Address & ":" & Range("tAccountName").Offset(zOpenTrades.Count, 0).Address), _
            Range("aAccountName").Offset(i, 0).Value, _
            Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address))
        
        'Performs an amount calculation using the open trades section
        Range("aSizeInUSD").Offset(i, 0).Value = Application.WorksheetFunction.SumIf( _
            Range(Range("tAccountName").Offset(1, 0).Address & ":" & Range("tAccountName").Offset(zOpenTrades.Count, 0).Address), _
            Range("aAccountName").Offset(i, 0).Value, _
            Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address))
    
        ' Performs an acccurate leverage calculation
        Range("aLeverage").Offset(i, 0).Value = Range("aSizeInUSD").Offset(i, 0).Value / Range("aEquity").Offset(i, 0).Value
    Next i
    Call RefreshTotalAccounts
End Sub

Private Sub ColorAccounts()
    If zAccounts.Count > 0 Then
        
        ' Clears any existing formatting
        Range(Range("aAccountName").Offset(1, 0).Address & ":" & Range("aBaseUnitSize").Offset(zAccounts.Count, 0).Address).Interior.Color = xlNone
        Range(Range("aAccountName").Offset(1, 0).Address & ":" & Range("aBaseUnitSize").Offset(zAccounts.Count, 0).Address).Select
        
        ' Colors every other row a different color for improved readability
        For i = 1 To Selection.Rows.Count
            If i Mod 2 = 1 Then
                Selection.Rows(i).Interior.Color = zColorAlternatingRow
            End If
        Next
    End If
    Range("iStart").Select
End Sub

Private Sub ShowOpenTrades()
    On Error GoTo Catch
    
    ' Sets section column heading value
    Range("tOpenTradesData") = "Open Trades Data (" & zOpenTrades.Count & ")"
    
    ' Inserts the number of rows necessary to display all open trades + a total row
    For i = 0 To zOpenTrades.Count
        Range("tAccountName").Offset(2, 0).EntireRow.Insert
    Next i
    
    ' Iterates through all open trades
    For i = 0 To zOpenTrades.Count - 1
        Set zOpenTradeRow = zOpenTrades.getRow(i)
        
        'Applies formatting
        Range("tAccountName").Offset(i + 1, 0).Resize(, zLengthOpenTrades).Font.Color = vbBlack
        Range("tAccountName").Offset(i + 1, 0).NumberFormat = zNumberFormatText
        Range("tPipCost").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("tAmount").Offset(i + 1, 0).NumberFormat = zNumberFormatComma
        Range("tSizeInUSD").Offset(i + 1, 0).NumberFormat = zNumberFormatUSD
        Range("tPL").Offset(i + 1, 0).NumberFormat = zNumberFormatPL
        Range("tGrossPL").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("tOpenTime").Offset(i + 1, 0).NumberFormat = zNumberFormatTime
        Range("tDurationOpen").Offset(i + 1, 0).HorizontalAlignment = xlRight
        Range("tRolloverInterest").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("tUsedMargin").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency

        ' Names cell for future reference
        Names.Add Name:="OpenTicket_" & zOpenTradeRow.TradeID, RefersTo:=Range("tTicketNumber").Offset(i + 1)
        
        ' Populates values
        Range("tTicketNumber").Offset(i + 1, 0).Value = zOpenTradeRow.TradeID
        Range("tAccountName").Offset(i + 1, 0).Value = zOpenTradeRow.AccountName
        Range("tAmount").Offset(i + 1, 0).Value = zOpenTradeRow.Amount
        Range("tPL").Offset(i + 1, 0).Value = zOpenTradeRow.PL
        Range("tGrossPL").Offset(i + 1, 0).Value = zOpenTradeRow.GrossPL
        Range("tOpenTime").Offset(i + 1, 0).Value = zOpenTradeRow.OpenTime
        Range("tOpenRate").Offset(i + 1, 0).Value = zOpenTradeRow.OpenRate
        Range("tClose").Offset(i + 1, 0).Value = zOpenTradeRow.Close
        Range("tRolloverInterest").Offset(i + 1, 0).Value = zOpenTradeRow.RolloverInterest
        Range("tUsedMargin").Offset(i + 1, 0).Value = zOpenTradeRow.UsedMargin
        
        ' Populates duration open
        zDifferenceMinutes = DateDiff("n", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) Mod 60
        If zDifferenceMinutes < 0 Then zDifferenceMinutes = 0
        If zDifferenceMinutes < 10 Then
            zDifferenceMinutesDisplay = "0" & Val(zDifferenceMinutes)
        Else
            zDifferenceMinutesDisplay = zDifferenceMinutes
        End If
        zDifferenceHours = DateDiff("h", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) Mod 24
        If zDifferenceHours < 10 Then
            zDifferenceHoursDisplay = "0" & Val(zDifferenceHours)
        Else
            zDifferenceHoursDisplay = zDifferenceHours
        End If
        zDifferenceDays = DateDiff("h", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) / 24
        If zDifferenceDays < 10 Then
            zDifferenceDaysDisplay = "0" & Val(zDifferenceDays)
        Else
            zDifferenceDaysDisplay = zDifferenceDays
        End If
        Range("tDurationOpen").Offset(i + 1, 0).Value = zDifferenceDaysDisplay & " d | " & _
            zDifferenceHoursDisplay & " h | " & zDifferenceMinutesDisplay & " m"

        ' Searches hidden offer table for data
        Range("tOfferID").Offset(i + 1, 0).Value = Range("SymbolIDForInstrument_" & zOpenTradeRow.OfferID)
        Range("tCounterCurrency").Offset(i + 1, 0).Value = Range("SymbolIDForCounterCurrency_" & zOpenTradeRow.OfferID)
        Range("tBaseCurrency").Offset(i + 1, 0).Value = Range("SymbolIDForContractCurrency_" & zOpenTradeRow.OfferID)
        Range("tSizeInUSD").Offset(i + 1, 0).Value = Range("SymbolIDForSizeInUSD_" & zOpenTradeRow.OfferID) * (zOpenTradeRow.Amount / 1000)
        
        ' Error handling for trades where PL = 0
        ' If 0 it cannot be calculated until the next price update
        If zOpenTradeRow.PL <> 0 Then
            Range("tPipCost").Offset(i + 1, 0).Value = (zOpenTradeRow.GrossPL / zOpenTradeRow.PL)
        End If
        
        ' Populates direction value and colors text
        If zOpenTradeRow.BuySell = "B" Then
            With Range("tBuySell").Offset(i + 1, 0)
                .Value = "Long"
                .Font.Color = zColorPositive
            End With
        Else
            With Range("tBuySell").Offset(i + 1, 0)
                .Value = "Short"
                .Font.Color = zColorNegative
            End With
        End If
        
        ' Colors p/l & gross p/l text
        If zOpenTradeRow.PL < 0 Then
            Range("tPL").Offset(i + 1, 0).Font.Color = zColorNegative
            Range("tGrossPL").Offset(i + 1, 0).Font.Color = zColorNegative
        ElseIf zOpenTradeRow.PL = 0 Then
            Range("tPL").Offset(i + 1, 0).Font.Color = zColorNeutral
            Range("tGrossPL").Offset(i + 1, 0).Font.Color = zColorNeutral
        Else
            Range("tPL").Offset(i + 1, 0).Font.Color = zColorPositive
            Range("tGrossPL").Offset(i + 1, 0).Font.Color = zColorPositive
        End If

        ' Colors rollover text
        If zOpenTradeRow.RolloverInterest < 0 Then
            Range("tRolloverInterest").Offset(i + 1, 0).Font.Color = zColorNegative
        ElseIf zOpenTradeRow.RolloverInterest = 0 Then
            Range("tRolloverInterest").Offset(i + 1, 0).Font.Color = zColorNeutral
        Else
            Range("tRolloverInterest").Offset(i + 1, 0).Font.Color = zColorPositive
        End If
    Next i
    
    ' Names total row cells for future reference and applies formatting
    Names.Add Name:="Total_tAccountName", RefersTo:=Range("tAccountName").Offset(zOpenTrades.Count + 1, 0)
    Names.Add Name:="Total_tAmount", RefersTo:=Range("tAmount").Offset(zOpenTrades.Count + 1, 0)
    Names.Add Name:="Total_tSizeInUSD", RefersTo:=Range("tSizeInUSD").Offset(zOpenTrades.Count + 1, 0)
    Names.Add Name:="Total_tPL", RefersTo:=Range("tPL").Offset(zOpenTrades.Count + 1, 0)
    Names.Add Name:="Total_tGrossPL", RefersTo:=Range("tGrossPL").Offset(zOpenTrades.Count + 1, 0)
    Names.Add Name:="Total_tRolloverInterest", RefersTo:=Range("tRolloverInterest").Offset(zOpenTrades.Count + 1, 0)
    Names.Add Name:="Total_tUsedMargin", RefersTo:=Range("tUsedMargin").Offset(zOpenTrades.Count + 1, 0)
    With Range("Total_tAccountName")
        .Resize(, zLengthOpenTrades).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Resize(, zLengthOpenTrades).Borders(xlEdgeTop).Color = zColorTotalLine
        .Resize(, zLengthOpenTrades).Interior.Color = zColorTotalRow
    End With
    Range("Total_tAmount").NumberFormat = zNumberFormatComma
    Range("Total_tSizeInUSD").NumberFormat = zNumberFormatUSD
    Range("Total_tPL").NumberFormat = zNumberFormatPL
    Range("Total_tGrossPL").NumberFormat = zNumberFormatCurrency
    Range("Total_tRolloverInterest").NumberFormat = zNumberFormatCurrency
    Range("Total_tUsedMargin").NumberFormat = zNumberFormatCurrency
    
    Call ColorOpenTrades
    Call RefreshTotalOpenTrades
    Call AccountsHelper
    Exit Sub
Catch:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error located in sub ShowOpenTrades()"
End Sub

Private Sub zOpenTrades_Changed(ByVal xRowID As String, ByVal xAccountRowMessage As fxcore2_com.IRow)
    On Error Resume Next
    Set zOpenTradeRow = xAccountRowMessage
    
    ' Determines where to place the updated values
    i = Range("OpenTicket_" & zOpenTradeRow.TradeID).Row - Range("tAccountName").Row
    
    ' Populates values
    Range("tTicketNumber").Offset(i, 0).Value = zOpenTradeRow.TradeID
    Range("tAmount").Offset(i, 0).Value = zOpenTradeRow.Amount
    Range("tPL").Offset(i, 0).Value = zOpenTradeRow.PL
    Range("tClose").Offset(i, 0).Value = zOpenTradeRow.Close
    Range("tGrossPL").Offset(i, 0).Value = zOpenTradeRow.GrossPL
    Range("tRolloverInterest").Offset(i, 0).Value = zOpenTradeRow.RolloverInterest
    Range("tUsedMargin").Offset(i, 0).Value = zOpenTradeRow.UsedMargin
    Range("tSizeInUSD").Offset(i, 0).Value = Range("SymbolIDForSizeInUSD_" & zOpenTradeRow.OfferID) * (zOpenTradeRow.Amount / 1000)
    
    ' Error handling for trades where PL = 0
    ' If 0 it cannot be calculated until the next price update
    If zOpenTradeRow.PL <> 0 Then
        Range("tPipCost").Offset(i, 0).Value = (zOpenTradeRow.GrossPL / zOpenTradeRow.PL)
    End If
    
    ' Colors p/l & gross p/l text
    If zOpenTradeRow.PL < 0 Then
        Range("tPL").Offset(i, 0).Font.Color = zColorNegative
        Range("tGrossPL").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zOpenTradeRow.PL = 0 Then
        Range("tPL").Offset(i, 0).Font.Color = zColorNeutral
        Range("tGrossPL").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("tPL").Offset(i, 0).Font.Color = zColorPositive
        Range("tGrossPL").Offset(i, 0).Font.Color = zColorPositive
    End If
    
    ' Colors rollover text
    If zOpenTradeRow.RolloverInterest < 0 Then
        Range("tRolloverInterest").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zOpenTradeRow.RolloverInterest = 0 Then
        Range("tRolloverInterest").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("tRolloverInterest").Offset(i, 0).Font.Color = zColorPositive
    End If
    
    Call RefreshOpenTradesDurationOpen
    Call RefreshTotalOpenTrades
    'Call CurrencyChanged
    Exit Sub
End Sub

Private Sub zOpenTrades_Added(ByVal xRowID As String, ByVal xOpenTradeRowMessage As fxcore2_com.IRow)
    On Error Resume Next

    ' Sets section column heading value
    Range("tOpenTradesData") = "Open Trades Data (" & zOpenTrades.Count & ")"

    ' Inserts a row
    Range("cClosedTradesData").Offset(-3, 0).EntireRow.Insert
    
    ' Names cell for future reference
    Set zOpenTradeRow = xOpenTradeRowMessage
    i = zOpenTrades.Count
    Names.Add Name:="OpenTicket_" & zOpenTradeRow.TradeID, RefersTo:=Range("tTicketNumber").Offset(i)
    
    'Applies formatting
    Range("tAccountName").Offset(i, 0).Resize(, zLengthOpenTrades).Font.Color = vbBlack
    Range("tAccountName").Offset(i, 0).NumberFormat = zNumberFormatText
    Range("tAmount").Offset(i, 0).NumberFormat = zNumberFormatComma
    Range("tPipCost").Offset(i, 0).NumberFormat = zNumberFormatCurrency
    Range("tSizeInUSD").Offset(i, 0).NumberFormat = zNumberFormatUSD
    Range("tPL").Offset(i, 0).NumberFormat = zNumberFormatPL
    Range("tGrossPL").Offset(i, 0).NumberFormat = zNumberFormatCurrency
    Range("tOpenTime").Offset(i, 0).NumberFormat = zNumberFormatTime
    Range("tDurationOpen").Offset(i, 0).HorizontalAlignment = xlRight
    Range("tRolloverInterest").Offset(i, 0).NumberFormat = zNumberFormatCurrency
    Range("tUsedMargin").Offset(i, 0).NumberFormat = zNumberFormatCurrency

    ' Populates values
    Range("tTicketNumber").Offset(i, 0).Value = zOpenTradeRow.TradeID
    Range("tAccountName").Offset(i, 0).Value = zOpenTradeRow.AccountName
    Range("tAmount").Offset(i, 0).Value = zOpenTradeRow.Amount
    Range("tPL").Offset(i, 0).Value = zOpenTradeRow.PL
    Range("tGrossPL").Offset(i, 0).Value = zOpenTradeRow.GrossPL
    Range("tOpenTime").Offset(i, 0).Value = zOpenTradeRow.OpenTime
    Range("tOpenRate").Offset(i, 0).Value = zOpenTradeRow.OpenRate
    Range("tClose").Offset(i, 0).Value = zOpenTradeRow.Close
    Range("tRolloverInterest").Offset(i, 0).Value = zOpenTradeRow.RolloverInterest
    Range("tUsedMargin").Offset(i, 0).Value = zOpenTradeRow.UsedMargin
    
    ' Populates duration open
    zDifferenceMinutes = DateDiff("n", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) Mod 60
    If zDifferenceMinutes < 0 Then zDifferenceMinutes = 0
    If zDifferenceMinutes < 10 Then
        zDifferenceMinutesDisplay = "0" & Val(zDifferenceMinutes)
    Else
        zDifferenceMinutesDisplay = zDifferenceMinutes
    End If
    zDifferenceHours = DateDiff("h", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) Mod 24
    If zDifferenceHours < 10 Then
        zDifferenceHoursDisplay = "0" & Val(zDifferenceHours)
    Else
        zDifferenceHoursDisplay = zDifferenceHours
    End If
    zDifferenceDays = DateDiff("h", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) / 24
    If zDifferenceDays < 10 Then
        zDifferenceDaysDisplay = "0" & Val(zDifferenceDays)
    Else
        zDifferenceDaysDisplay = zDifferenceDays
    End If
    Range("tDurationOpen").Offset(i, 0).Value = zDifferenceDaysDisplay & " d | " & _
        zDifferenceHoursDisplay & " h | " & zDifferenceMinutesDisplay & " m"

    ' Searches hidden offer table for data
    Range("tOfferID").Offset(i, 0).Value = Range("SymbolIDForInstrument_" & zOpenTradeRow.OfferID)
    Range("tBaseCurrency").Offset(i, 0).Value = Range("SymbolIDForContractCurrency_" & zOpenTradeRow.OfferID)
    Range("tCounterCurrency").Offset(i, 0).Value = Range("SymbolIDForCounterCurrency_" & zOpenTradeRow.OfferID)
    Range("tSizeInUSD").Offset(i, 0).Value = Range("SymbolIDForSizeInUSD_" & zOpenTradeRow.OfferID) * (zOpenTradeRow.Amount / 1000)
    
    ' Error handling for trades where PL = 0
    If zOpenTradeRow.PL <> 0 Then
        Range("tPipCost").Offset(i, 0).Value = (zOpenTradeRow.GrossPL / zOpenTradeRow.PL)
    End If

    ' Colors direction text
    If zOpenTradeRow.BuySell = "B" Then
        With Range("tBuySell").Offset(i, 0)
            .Value = "Long"
            .Font.Color = zColorPositive
        End With
    Else
        With Range("tBuySell").Offset(i, 0)
            .Value = "Short"
            .Font.Color = zColorNegative
        End With
    End If
    
    ' Colors p/l & gross p/l text
    If zOpenTradeRow.PL < 0 Then
        Range("tPL").Offset(i, 0).Font.Color = zColorNegative
        Range("tGrossPL").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zOpenTradeRow.PL = 0 Then
        Range("tPL").Offset(i, 0).Font.Color = zColorNeutral
        Range("tGrossPL").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("tPL").Offset(i, 0).Font.Color = zColorPositive
        Range("tGrossPL").Offset(i, 0).Font.Color = zColorPositive
    End If

    ' Colors rollover text
    If zOpenTradeRow.RolloverInterest < 0 Then
        Range("tRolloverInterest").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zOpenTradeRow.RolloverInterest = 0 Then
        Range("tRolloverInterest").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("tRolloverInterest").Offset(i, 0).Font.Color = zColorPositive
    End If
    
    CreateSound (SymbolToSpeakableName(zOpenTradeRow.OfferID) & zTextToSpeakOpenTrades)
    
    Call RefreshOpenTradesDurationOpen
    Call ReSortOpenTrades
    Call ColorOpenTrades
    Call RefreshTotalOpenTrades
    Call ClearCurrencySummary
    Call ShowCurrencySummary
End Sub

Private Sub zOpenTrades_Deleted(ByVal rowID As String, ByVal xOpenTradeRowMessage As fxcore2_com.IRow)
    On Error Resume Next
    
    ' Sets section column heading value
    Range("tOpenTradesData") = "Open Trades Data (" & zOpenTrades.Count & ")"
    
    Set zOpenTradeRow = xOpenTradeRowMessage
    i = zOpenTrades.Count
    
    Range("OpenTicket_" & zOpenTradeRow.TradeID).EntireRow.Delete
    For Each zNamedCell In Names
        If InStr(1, zNamedCell.RefersTo, "#REF!") > 0 Then zNamedCell.Delete
    Next zNamedCell
    
    Call RefreshOpenTradesDurationOpen
    Call ColorOpenTrades
    Call RefreshTotalOpenTrades
    Call ClearCurrencySummary
    Call ShowCurrencySummary
End Sub

Private Sub RefreshOpenTradesDurationOpen()
    On Error Resume Next
    
    ' Limits code to only run every 5 minutes; done to minimize cpu load
    If (Now() - zLastRefreshTime) > TimeValue("00:05:00") Then
        For i = 0 To zOpenTrades.Count - 1
            Set zOpenTradeRow = zOpenTrades.getRow(i)
            k = (Range("OpenTicket_" & zOpenTradeRow.TradeID).Row - Range("tAccountName").Row)
            
            ' Populates duration open
            zDifferenceMinutes = DateDiff("n", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) Mod 60
            If zDifferenceMinutes < 0 Then zDifferenceMinutes = 0
            If zDifferenceMinutes < 10 Then
                zDifferenceMinutesDisplay = "0" & Val(zDifferenceMinutes)
            Else
                zDifferenceMinutesDisplay = zDifferenceMinutes
            End If
            zDifferenceHours = DateDiff("h", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) Mod 24
            If zDifferenceHours < 10 Then
                zDifferenceHoursDisplay = "0" & Val(zDifferenceHours)
            Else
                zDifferenceHoursDisplay = zDifferenceHours
            End If
            zDifferenceDays = DateDiff("h", zOpenTradeRow.OpenTime, DateAdd("h", zUTCOffset, Now())) / 24
            If zDifferenceDays < 10 Then
                zDifferenceDaysDisplay = "0" & Val(zDifferenceDays)
            Else
                zDifferenceDaysDisplay = zDifferenceDays
            End If
            Range("tDurationOpen").Offset(k, 0).Value = zDifferenceDaysDisplay & " d | " & _
                zDifferenceHoursDisplay & " h | " & zDifferenceMinutesDisplay & " m"
        Next i
        zLastRefreshTime = Now()
    End If
End Sub

Private Sub RefreshTotalOpenTrades()
    On Error Resume Next
    If zOpenTrades.Count > 0 Then
        ' Populates Values
        Range("Total_tAmount").Value = Application.WorksheetFunction.Sum( _
            Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address))
        Range("Total_tSizeInUSD").Value = Application.WorksheetFunction.Sum( _
            Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address))
        Range("Total_tUsedMargin").Value = Application.WorksheetFunction.Sum( _
            Range(Range("tUsedMargin").Offset(1, 0).Address & ":" & Range("tUsedMargin").Offset(zOpenTrades.Count, 0).Address))
            
        ' Calculates and colors total p/l
        Range("Total_tPL").Value = Application.WorksheetFunction.Sum( _
            Range(Range("tPL").Offset(1, 0).Address & ":" & Range("tPL").Offset(zOpenTrades.Count, 0).Address))
        If Range("Total_tPL").Value < 0 Then
            Range("Total_tPL").Font.Color = zColorNegative
        ElseIf Range("Total_tPL").Value = 0 Then
            Range("Total_tPL").Font.Color = zColorNeutral
        Else
            Range("Total_tPL").Font.Color = zColorPositive
        End If
        
        ' Calculates and colors total gross p/l
        Range("Total_tGrossPL").Value = Application.WorksheetFunction.Sum( _
            Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address))
        If Range("Total_tGrossPL").Value < 0 Then
            Range("Total_tGrossPL").Font.Color = zColorNegative
        ElseIf Range("Total_tGrossPL").Value = 0 Then
            Range("Total_tGrossPL").Font.Color = zColorNeutral
        Else
            Range("Total_tGrossPL").Font.Color = zColorPositive
        End If
        
        ' Calculates and colors total rollover
        Range("Total_tRolloverInterest").Value = Application.WorksheetFunction.Sum( _
            Range(Range("tRolloverInterest").Offset(1, 0).Address & ":" & Range("tRolloverInterest").Offset(zOpenTrades.Count, 0).Address))
        If Range("Total_tRolloverInterest").Value < 0 Then
            Range("Total_tRolloverInterest").Font.Color = zColorNegative
        ElseIf Range("Total_tRolloverInterest").Value = 0 Then
            Range("Total_tRolloverInterest").Font.Color = zColorNeutral
        Else
            Range("Total_tRolloverInterest").Font.Color = zColorPositive
        End If
    Else
        ' If there are no open trades, this clears any existing formatting and sets values to 0 in the total row
        Range("Total_tAmount").Value = vbNullString
        Range("Total_tSizeInUSD").Value = vbNullString
        Range("Total_tPL").Value = vbNullString
        Range("Total_tGrossPL").Value = vbNullString
        Range("Total_tRolloverInterest").Value = vbNullString
        Range("Total_tUsedMargin").Value = vbNullString
        Range("Total_tAccountName").Resize(, zLengthOpenTrades).Interior.Color = xlNone
    End If
End Sub

Private Sub ReSortOpenTrades()
    For Each zCell In Range("tAccountName").Resize(, zLengthOpenTrades)
        If zCell Like "*" & "(" & "*" Then
            For Each zNamedCell In Names
                If InStr(1, zNamedCell.Name, "OpenTicket_") > 0 Then zNamedCell.Delete
            Next zNamedCell
            If zOpenTradesSortStatus = "Descending" Then
                ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                    Key:=Range(zCell.Offset(1, 0).Address & ":" & zCell.Offset(zOpenTrades.Count, 0).Address), _
                    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            Else
                ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                    Key:=Range(zCell.Offset(1, 0).Address & ":" & zCell.Offset(zOpenTrades.Count, 0).Address), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            End If
            With ActiveWorkbook.Worksheets("Account Dashboard").Sort
                .SetRange Range(Range("tTicketNumber").Address & ":" & Range("tUsedMargin").Offset(zOpenTrades.Count, 0).Address)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            For i = 0 To zOpenTrades.Count - 1
                Names.Add Name:="OpenTicket_" & Range("tTicketNumber").Offset(i + 1), RefersTo:=Range("tTicketNumber").Offset(i + 1)
            Next i
        End If
    Next zCell
End Sub

Private Sub ColorOpenTrades()
    If zOpenTrades.Count > 0 Then
    
        ' Clears any existing formatting
        Range(Range("tAccountName").Offset(1, 0).Address & ":" & Range("tUsedMargin").Offset(zOpenTrades.Count, 0).Address).Interior.Color = xlNone
        Range(Range("tAccountName").Offset(1, 0).Address & ":" & Range("tUsedMargin").Offset(zOpenTrades.Count, 0).Address).Select
                
         ' Colors every other row a different color for improved readability
        For i = 1 To Selection.Rows.Count
            If i Mod 2 = 1 Then
                Selection.Rows(i).Interior.Color = zColorAlternatingRow
            End If
        Next
    End If
    Range("iStart").Select
End Sub

Private Sub ShowClosedTrades()
    On Error GoTo Catch
    
    ' Sets section column heading value
    Range("cClosedTradesData") = "Closed Trades Data (" & zClosedTrades.Count & ")"
    
    ' Inserts the number of rows necessary to display all closed trades + a total row
    For i = 0 To zClosedTrades.Count
        Range("cAccountName").Offset(2, 0).EntireRow.Insert
    Next i
    
    ' Iterates through all closed trades
    For i = 0 To zClosedTrades.Count - 1
        Set zClosedTradeRow = zClosedTrades.getRow(i)
        
        'Applies formatting
        Range("cAccountName").Offset(i + 1, 0).Resize(, zLengthClosedTrades).Font.Color = vbBlack
        Range("cAccountName").Offset(i + 1, 0).NumberFormat = zNumberFormatText
        Range("cAmount").Offset(i + 1, 0).NumberFormat = zNumberFormatComma
        Range("cPipCost").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("cSizeInUSD").Offset(i + 1, 0).NumberFormat = zNumberFormatUSD
        Range("cPL").Offset(i + 1, 0).NumberFormat = zNumberFormatPL
        Range("cGrossPL").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        Range("cOpenTime").Offset(i + 1, 0).NumberFormat = zNumberFormatTime
        Range("cDurationOpen").Offset(i + 1, 0).HorizontalAlignment = xlRight
        Range("cCloseTime").Offset(i + 1, 0).NumberFormat = zNumberFormatTime
        Range("cRolloverInterest").Offset(i + 1, 0).NumberFormat = zNumberFormatCurrency
        
        ' Names cell for future reference
        Names.Add Name:="ClosedTicket_" & zClosedTradeRow.TradeID, RefersTo:=Range("cTicketNumber").Offset(i + 1)
        
        ' Populates values
        Range("cTicketNumber").Offset(i + 1, 0).Value = zClosedTradeRow.TradeID
        Range("cAccountName").Offset(i + 1, 0).Value = zClosedTradeRow.AccountName
        Range("cAmount").Offset(i + 1, 0).Value = zClosedTradeRow.Amount
        Range("cPL").Offset(i + 1, 0).Value = zClosedTradeRow.PL
        Range("cGrossPL").Offset(i + 1, 0).Value = zClosedTradeRow.GrossPL
        Range("cOpenTime").Offset(i + 1, 0).Value = zClosedTradeRow.OpenTime
        Range("cOpenRate").Offset(i + 1, 0).Value = zClosedTradeRow.OpenRate
        Range("cCloseTime").Offset(i + 1, 0).Value = zClosedTradeRow.CloseTime
        Range("cCloseRate").Offset(i + 1, 0).Value = zClosedTradeRow.CloseRate
        Range("cRolloverInterest").Offset(i + 1, 0).Value = zClosedTradeRow.RolloverInterest
        
        ' Populates duration open
        zDifferenceMinutes = DateDiff("n", zClosedTradeRow.OpenTime, zClosedTradeRow.CloseTime) Mod 60
        If zDifferenceMinutes < 0 Then zDifferenceMinutes = 0
        If zDifferenceMinutes < 10 Then
            zDifferenceMinutesDisplay = "0" & Val(zDifferenceMinutes)
        Else
            zDifferenceMinutesDisplay = zDifferenceMinutes
        End If
        zDifferenceHours = DateDiff("h", zClosedTradeRow.OpenTime, zClosedTradeRow.CloseTime) Mod 24
        If zDifferenceHours < 10 Then
            zDifferenceHoursDisplay = "0" & Val(zDifferenceHours)
        Else
            zDifferenceHoursDisplay = zDifferenceHours
        End If
        zDifferenceDays = DateDiff("h", zClosedTradeRow.OpenTime, zClosedTradeRow.CloseTime) / 24
        If zDifferenceDays < 10 Then
            zDifferenceDaysDisplay = "0" & Val(zDifferenceDays)
        Else
            zDifferenceDaysDisplay = zDifferenceDays
        End If
        Range("cDurationOpen").Offset(i + 1, 0).Value = zDifferenceDaysDisplay & " d | " & _
            zDifferenceHoursDisplay & " h | " & zDifferenceMinutesDisplay & " m"

        ' Searches hidden offer table for data
        Range("cOfferID").Offset(i + 1, 0).Value = Range("SymbolIDForInstrument_" & zClosedTradeRow.OfferID)
        Range("cCounterCurrency").Offset(i + 1, 0).Value = Range("SymbolIDForCounterCurrency_" & zClosedTradeRow.OfferID)
        Range("cBaseCurrency").Offset(i + 1, 0).Value = Range("SymbolIDForContractCurrency_" & zClosedTradeRow.OfferID)
        Range("cSizeInUSD").Offset(i + 1, 0).Value = Range("SymbolIDForSizeInUSD_" & zClosedTradeRow.OfferID) * (zClosedTradeRow.Amount / 1000)
        
        ' Error handling for trades where PL = 0; if 0 it cannot be calculated
        If zClosedTradeRow.PL <> 0 Then
            Range("cPipCost").Offset(i + 1, 0).Value = (zClosedTradeRow.GrossPL / zClosedTradeRow.PL)
        End If
        
        ' Populates direction value and colors text
        If zClosedTradeRow.BuySell = "B" Then
            With Range("cBuySell").Offset(i + 1, 0)
                .Value = "Long"
                .Font.Color = zColorPositive
            End With
        Else
            With Range("cBuySell").Offset(i + 1, 0)
                .Value = "Short"
                .Font.Color = zColorNegative
            End With
        End If
        
        ' Colors p/l & gross p/l text
        If zClosedTradeRow.PL < 0 Then
            Range("cPL").Offset(i + 1, 0).Font.Color = zColorNegative
            Range("cGrossPL").Offset(i + 1, 0).Font.Color = zColorNegative
        ElseIf zClosedTradeRow.PL = 0 Then
            Range("cPL").Offset(i + 1, 0).Font.Color = zColorNeutral
            Range("cGrossPL").Offset(i + 1, 0).Font.Color = zColorNeutral
        Else
            Range("cPL").Offset(i + 1, 0).Font.Color = zColorPositive
            Range("cGrossPL").Offset(i + 1, 0).Font.Color = zColorPositive
        End If
        
        ' Colors rollover text
        If zClosedTradeRow.RolloverInterest < 0 Then
            Range("cRolloverInterest").Offset(i + 1, 0).Font.Color = zColorNegative
        ElseIf zClosedTradeRow.RolloverInterest = 0 Then
            Range("cRolloverInterest").Offset(i + 1, 0).Font.Color = zColorNeutral
        Else
            Range("cRolloverInterest").Offset(i + 1, 0).Font.Color = zColorPositive
        End If
    Next i
    
    ' Names total row cells for future reference and applies formatting
    Names.Add Name:="Total_cAccountName", RefersTo:=Range("cAccountName").Offset(zClosedTrades.Count + 1, 0)
    Names.Add Name:="Total_cAmount", RefersTo:=Range("cAmount").Offset(zClosedTrades.Count + 1, 0)
    Names.Add Name:="Total_cSizeInUSD", RefersTo:=Range("cSizeInUSD").Offset(zClosedTrades.Count + 1, 0)
    Names.Add Name:="Total_cPL", RefersTo:=Range("cPL").Offset(zClosedTrades.Count + 1, 0)
    Names.Add Name:="Total_cGrossPL", RefersTo:=Range("cGrossPL").Offset(zClosedTrades.Count + 1, 0)
    Names.Add Name:="Total_cRolloverInterest", RefersTo:=Range("cRolloverInterest").Offset(zClosedTrades.Count + 1, 0)
    With Range("Total_cAccountName")
        .Resize(, zLengthClosedTrades).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Resize(, zLengthClosedTrades).Borders(xlEdgeTop).Color = zColorTotalLine
        .Resize(, zLengthClosedTrades).Interior.Color = zColorTotalRow
    End With
    Range("Total_cAmount").NumberFormat = zNumberFormatComma
    Range("Total_cSizeInUSD").NumberFormat = zNumberFormatUSD
    Range("Total_cPL").NumberFormat = zNumberFormatPL
    Range("Total_cGrossPL").NumberFormat = zNumberFormatCurrency
    Range("Total_cRolloverInterest").NumberFormat = zNumberFormatCurrency
    
    Call ColorClosedTrades
    Call RefreshTotalClosedTrades
    Exit Sub
Catch:
    MsgBox Err.Description & vbCrLf & vbCrLf & "Error located in sub ShowClosedTrades()"
End Sub

Private Sub zClosedTrades_Added(ByVal xRowID As String, ByVal xClosedTradeRowMessage As fxcore2_com.IRow)
    On Error Resume Next
    
    ' Sets section column heading value
    Range("cClosedTradesData") = "Closed Trades Data (" & zClosedTrades.Count & ")"
    
    ' Inserts a row
    Range("uCurrencyData").Offset(-3, 0).EntireRow.Insert
    
    ' Names cell for future reference
    Set zClosedTradeRow = xClosedTradeRowMessage
    i = zClosedTrades.Count
    Names.Add Name:="ClosedTicket_" & zClosedTradeRow.TradeID, RefersTo:=Range("cTicketNumber").Offset(i)
    
    'Applies formatting
    Range("cAccountName").Offset(i, 0).Resize(, zLengthClosedTrades).Font.Color = vbBlack
    Range("cAccountName").Offset(i, 0).NumberFormat = zNumberFormatText
    Range("cAmount").Offset(i, 0).NumberFormat = zNumberFormatComma
    Range("cPipCost").Offset(i, 0).NumberFormat = zNumberFormatCurrency
    Range("cSizeInUSD").Offset(i, 0).NumberFormat = zNumberFormatUSD
    Range("cPL").Offset(i, 0).NumberFormat = zNumberFormatPL
    Range("cGrossPL").Offset(i, 0).NumberFormat = zNumberFormatCurrency
    Range("cOpenTime").Offset(i, 0).NumberFormat = zNumberFormatTime
    Range("cDurationOpen").Offset(i, 0).HorizontalAlignment = xlRight
    Range("cCloseTime").Offset(i, 0).NumberFormat = zNumberFormatTime
    Range("cRolloverInterest").Offset(i, 0).NumberFormat = zNumberFormatCurrency
    
    ' Populates values
    Range("cTicketNumber").Offset(i, 0).Value = zClosedTradeRow.TradeID
    Range("cAccountName").Offset(i, 0).Value = zClosedTradeRow.AccountName
    Range("cAmount").Offset(i, 0).Value = zClosedTradeRow.Amount
    Range("cPL").Offset(i, 0).Value = zClosedTradeRow.PL
    Range("cGrossPL").Offset(i, 0).Value = zClosedTradeRow.GrossPL
    Range("cOpenTime").Offset(i, 0).Value = zClosedTradeRow.OpenTime
    Range("cOpenRate").Offset(i, 0).Value = zClosedTradeRow.OpenRate
    Range("cCloseTime").Offset(i, 0).Value = zClosedTradeRow.CloseTime
    Range("cCloseRate").Offset(i, 0).Value = zClosedTradeRow.CloseRate
    Range("cRolloverInterest").Offset(i, 0).Value = zClosedTradeRow.RolloverInterest
    
    ' Populates duration open
    zDifferenceMinutes = DateDiff("n", zClosedTradeRow.OpenTime, zClosedTradeRow.CloseTime) Mod 60
    If zDifferenceMinutes < 0 Then zDifferenceMinutes = 0
    If zDifferenceMinutes < 10 Then
        zDifferenceMinutesDisplay = "0" & Val(zDifferenceMinutes)
    Else
        zDifferenceMinutesDisplay = zDifferenceMinutes
    End If
    zDifferenceHours = DateDiff("h", zClosedTradeRow.OpenTime, zClosedTradeRow.CloseTime) Mod 24
    If zDifferenceHours < 10 Then
        zDifferenceHoursDisplay = "0" & Val(zDifferenceHours)
    Else
        zDifferenceHoursDisplay = zDifferenceHours
    End If
    zDifferenceDays = DateDiff("h", zClosedTradeRow.OpenTime, zClosedTradeRow.CloseTime) / 24
    If zDifferenceDays < 10 Then
        zDifferenceDaysDisplay = "0" & Val(zDifferenceDays)
    Else
        zDifferenceDaysDisplay = zDifferenceDays
    End If
    Range("cDurationOpen").Offset(i, 0).Value = zDifferenceDaysDisplay & " d | " & _
        zDifferenceHoursDisplay & " h | " & zDifferenceMinutesDisplay & " m"

    ' Searches hidden offer table for data
    Range("cOfferID").Offset(i, 0).Value = Range("SymbolIDForInstrument_" & zClosedTradeRow.OfferID)
    Range("cBaseCurrency").Offset(i, 0).Value = Range("SymbolIDForContractCurrency_" & zClosedTradeRow.OfferID)
    Range("cCounterCurrency").Offset(i, 0).Value = Range("SymbolIDForCounterCurrency_" & zClosedTradeRow.OfferID)
    Range("cSizeInUSD").Offset(i, 0).Value = Range("SymbolIDForSizeInUSD_" & zClosedTradeRow.OfferID) * (zClosedTradeRow.Amount / 1000)
    
    ' Error handling for trades where PL = 0; if 0 it cannot be calculated
    If zClosedTradeRow.PL <> 0 Then
        Range("cPipCost").Offset(i, 0).Value = (zClosedTradeRow.GrossPL / zClosedTradeRow.PL)
    End If

    ' Colors direction text
    If zClosedTradeRow.BuySell = "B" Then
        With Range("cBuySell").Offset(i, 0)
            .Value = "Long"
            .Font.Color = zColorPositive
        End With
    Else
        With Range("cBuySell").Offset(i, 0)
            .Value = "Short"
            .Font.Color = zColorNegative
        End With
    End If
    
    ' Colors p/l & gross p/l text
    If zClosedTradeRow.PL < 0 Then
        Range("cPL").Offset(i, 0).Font.Color = zColorNegative
        Range("cGrossPL").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zClosedTradeRow.PL = 0 Then
        Range("cPL").Offset(i, 0).Font.Color = zColorNeutral
        Range("cGrossPL").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("cPL").Offset(i, 0).Font.Color = zColorPositive
        Range("cGrossPL").Offset(i, 0).Font.Color = zColorPositive
    End If
    
    ' Colors rollover text
    If zClosedTradeRow.RolloverInterest < 0 Then
        Range("cRolloverInterest").Offset(i, 0).Font.Color = zColorNegative
    ElseIf zClosedTradeRow.RolloverInterest = 0 Then
        Range("cRolloverInterest").Offset(i, 0).Font.Color = zColorNeutral
    Else
        Range("cRolloverInterest").Offset(i, 0).Font.Color = zColorPositive
    End If
    
    CreateSound (SymbolToSpeakableName(zClosedTradeRow.OfferID) & zTextToSpeakClosedTrades)
    
    Call ReSortClosedTrades
    Call ColorClosedTrades
    Call RefreshTotalClosedTrades
End Sub

Private Sub RefreshTotalClosedTrades()
    On Error Resume Next
    If zClosedTrades.Count > 0 Then
        ' Populates Values
        Range("Total_cAmount").Value = Application.WorksheetFunction.Sum( _
            Range(Range("cAmount").Offset(1, 0).Address & ":" & Range("cAmount").Offset(zClosedTrades.Count, 0).Address))
        Range("Total_cSizeInUSD").Value = Application.WorksheetFunction.Sum( _
            Range(Range("cSizeInUSD").Offset(1, 0).Address & ":" & Range("cSizeInUSD").Offset(zClosedTrades.Count, 0).Address))
        
        ' Calculates and colors total p/l
        Range("Total_cPL").Value = Application.WorksheetFunction.Sum( _
            Range(Range("cPL").Offset(1, 0).Address & ":" & Range("cPL").Offset(zClosedTrades.Count, 0).Address))
        If Range("Total_cPL").Value < 0 Then
            Range("Total_cPL").Font.Color = zColorNegative
        ElseIf Range("Total_cPL").Value = 0 Then
            Range("Total_cPL").Font.Color = zColorNeutral
        Else
            Range("Total_cPL").Font.Color = zColorPositive
        End If
        
        ' Calculates and colors total gross p/l
        Range("Total_cGrossPL").Value = Application.WorksheetFunction.Sum( _
            Range(Range("cGrossPL").Offset(1, 0).Address & ":" & Range("cGrossPL").Offset(zClosedTrades.Count, 0).Address))
        If Range("Total_cGrossPL").Value < 0 Then
            Range("Total_cGrossPL").Font.Color = zColorNegative
        ElseIf Range("Total_cGrossPL").Value = 0 Then
            Range("Total_cGrossPL").Font.Color = zColorNeutral
        Else
            Range("Total_cGrossPL").Font.Color = zColorPositive
        End If
        
        ' Calculates and colors total rollover
        Range("Total_cRolloverInterest").Value = Application.WorksheetFunction.Sum( _
            Range(Range("cRolloverInterest").Offset(1, 0).Address & ":" & Range("cRolloverInterest").Offset(zClosedTrades.Count, 0).Address))
        If Range("Total_cRolloverInterest").Value < 0 Then
            Range("Total_cRolloverInterest").Font.Color = zColorNegative
        ElseIf Range("Total_cRolloverInterest").Value = 0 Then
            Range("Total_cRolloverInterest").Font.Color = zColorNeutral
        Else
            Range("Total_cRolloverInterest").Font.Color = zColorPositive
        End If
    Else
        ' If there are no open trades, this clears any existing formatting and sets values to 0 in the total row
        Range("Total_cAmount").Value = vbNullString
        Range("Total_cSizeInUSD").Value = vbNullString
        Range("Total_cPL").Value = vbNullString
        Range("Total_cGrossPL").Value = vbNullString
        Range("Total_cRolloverInterest").Value = vbNullString
        Range("Total_cAccountName").Resize(, zLengthClosedTrades).Interior.Color = xlNone
    End If
End Sub

Private Sub ColorClosedTrades()
    If zClosedTrades.Count > 0 Then
        
        ' Clears any existing formatting
        Range(Range("cAccountName").Offset(1, 0).Address & ":" & Range("cRolloverInterest").Offset(zClosedTrades.Count, 0).Address).Interior.Color = xlNone
        Range(Range("cAccountName").Offset(1, 0).Address & ":" & Range("cRolloverInterest").Offset(zClosedTrades.Count, 0).Address).Select
        
        ' Colors every other row a different color for improved readability
        For i = 1 To Selection.Rows.Count
            If i Mod 2 = 1 Then
                Selection.Rows(i).Interior.Color = zColorAlternatingRow
            End If
        Next
    End If
    Range("iStart").Select
End Sub

Private Sub ReSortClosedTrades()
    For Each zCell In Range("cAccountName").Resize(, zLengthClosedTrades)
        If zCell Like "*" & "(" & "*" Then
            For Each zNamedCell In Names
                If InStr(1, zNamedCell.Name, "ClosedTicket_") > 0 Then zNamedCell.Delete
            Next zNamedCell
            If zOpenTradesSortStatus = "Descending" Then
                ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                    Key:=Range(zCell.Offset(1, 0).Address & ":" & zCell.Offset(zClosedTrades.Count, 0).Address), _
                    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            Else
                ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                    Key:=Range(zCell.Offset(1, 0).Address & ":" & zCell.Offset(zClosedTrades.Count, 0).Address), _
                    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            End If
            With ActiveWorkbook.Worksheets("Account Dashboard").Sort
                .SetRange Range(Range("cTicketNumber").Address & ":" & Range("cRolloverInterest").Offset(zClosedTrades.Count, 0).Address)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            For i = 0 To zClosedTrades.Count - 1
                Names.Add Name:="ClosedTicket_" & Range("cTicketNumber").Offset(i + 1), RefersTo:=Range("cTicketNumber").Offset(i + 1)
            Next i
        End If
    Next zCell
End Sub

Private Sub ShowCurrencySummary()
    On Error Resume Next
    
    ' This sub produces a runtime error, need to review!
    ' This sub produces a runtime error, need to review!
    ' This sub produces a runtime error, need to review!
    
    If zOpenTrades.Count > 0 Then
        Set zCurrencyCollection = CreateObject("Scripting.Dictionary")
        
        ' Calculates unique base curreny names
        If zOpenTrades.Count > 1 Then
            zCurrencyRange = Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address)
        Else
            zCurrencyRange = Range("tBaseCurrency").Offset(1, 0).Address
        End If
        For i = 1 To UBound(zCurrencyRange, 1)
            zCurrencyCollection(zCurrencyRange(i, 1)) = 1
        Next i
        
        ' Calculates unique counter curreny names
        If zOpenTrades.Count > 1 Then
            zCurrencyRange = Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address)
        Else
            zCurrencyRange = Range("tCounterCurrency").Offset(1, 0).Address
        End If
        For i = 1 To UBound(zCurrencyRange, 1)
            zCurrencyCollection(zCurrencyRange(i, 1)) = 1
        Next i
            
        ' Inserts the number of rows necessary
        For i = 1 To zCurrencyCollection.Count
            Range("uAccountName").Offset(2, 0).EntireRow.Insert
                        
            ' Applies formatting
            Range("uCurrency").Offset(i, 0).NumberFormat = zNumberFormatText
            Range("uLongAmount").Offset(i, 0).NumberFormat = zNumberFormatUSD
            Range("uLongSizeInUSD").Offset(i, 0).NumberFormat = zNumberFormatUSD
            Range("uLongGrossPL").Offset(i, 0).NumberFormat = zNumberFormatPL
            Range("uShortAmount").Offset(i, 0).NumberFormat = zNumberFormatUSD
            Range("uShortSizeInUSD").Offset(i, 0).NumberFormat = zNumberFormatUSD
            Range("uShortGrossPL").Offset(i, 0).NumberFormat = zNumberFormatPL
            Range("uBuySell").Offset(i, 0).NumberFormat = zNumberFormatText
            Range("uNetAmount").Offset(i, 0).NumberFormat = zNumberFormatUSD
            Range("uNetSizeInUSD").Offset(i, 0).NumberFormat = zNumberFormatUSD
            Range("uNetGrossPL").Offset(i, 0).NumberFormat = zNumberFormatPL
        Next i
        
        If zOpenTrades.Count > 1 Then
            Range("uCurrency").Offset(1, 0).Resize(zCurrencyCollection.Count) = Application.Transpose(zCurrencyCollection.Keys)
        Else
            Range("uCurrency").Offset(1, 0).Value = Range("tBaseCurrency").Offset(1, 0).Value
            Range("uCurrency").Offset(2, 0).Value = Range("tCounterCurrency").Offset(1, 0).Value
        End If
        
        For i = 1 To zCurrencyCollection.Count
            
            ' Names cells for future reference and applies formatting
            Names.Add Name:="CurrencySummary_" & Range("uCurrency").Offset(i, 0).Value, RefersTo:=Range("uCurrencyTicket").Offset(i, 0)
            
            ' Populates values
            Range("uLongAmount").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short")
            Range("uLongSizeInUSD").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short")
            Range("uLongGrossPL").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short")
            With Range("uLongGrossPL").Offset(i, 0)
                If .Value < 0 Then
                    .Font.Color = zColorNegative
                ElseIf .Value = 0 Then
                    .Font.Color = zColorNeutral
                Else
                    .Font.Color = zColorPositive
                End If
            End With
            
            Range("uShortAmount").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long")
            Range("uShortSizeInUSD").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long")
            Range("uShortGrossPL").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long")
            With Range("uShortGrossPL").Offset(i, 0)
                If .Value < 0 Then
                    .Font.Color = zColorNegative
                ElseIf .Value = 0 Then
                    .Font.Color = zColorNeutral
                Else
                    .Font.Color = zColorPositive
                End If
            End With
            
            If Range("uLongAmount").Offset(i, 0).Value = Range("uShortAmount").Offset(i, 0).Value Then
                Range("uBuySell").Offset(i, 0).Value = "Flat"
                Range("uBuySell").Offset(i, 0).Font.Color = zColorNeutral
            ElseIf Range("uLongAmount").Offset(i, 0).Value > Range("uShortAmount").Offset(i, 0).Value Then
                Range("uBuySell").Offset(i, 0).Value = "Long"
                Range("uBuySell").Offset(i, 0).Font.Color = zColorPositive
            Else
                Range("uBuySell").Offset(i, 0).Value = "Short"
                Range("uBuySell").Offset(i, 0).Font.Color = zColorNegative
            End If
            Range("uNetAmount").Offset(i, 0).Value = Range("uLongAmount").Offset(i, 0).Value - Range("uShortAmount").Offset(i, 0).Value
            Range("uNetSizeInUSD").Offset(i, 0).Value = Range("uLongSizeInUSD").Offset(i, 0).Value - Range("uShortSizeInUSD").Offset(i, 0).Value
            Range("uNetGrossPL").Offset(i, 0).Value = Range("uLongGrossPL").Offset(i, 0).Value - Range("uShortGrossPL").Offset(i, 0).Value
            With Range("uNetGrossPL").Offset(i, 0)
                If .Value < 0 Then
                    .Font.Color = zColorNegative
                ElseIf .Value = 0 Then
                    .Font.Color = zColorNeutral
                Else
                    .Font.Color = zColorPositive
                End If
            End With
        Next i

        Call ColorCurrencySummary
    End If
End Sub

Private Sub CurrencyChanged()
    On Error Resume Next
    If zOpenTrades.Count > 0 Then
        For i = 1 To zCurrencyCollection.Count
            
            ' Populates values
            Range("uLongAmount").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short")
            Range("uLongSizeInUSD").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short")
            Range("uLongGrossPL").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short")
            With Range("uLongGrossPL").Offset(i, 0)
                If .Value < 0 Then
                    .Font.Color = zColorNegative
                ElseIf .Value = 0 Then
                    .Font.Color = zColorNeutral
                Else
                    .Font.Color = zColorPositive
                End If
            End With
            
            Range("uShortAmount").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tAmount").Offset(1, 0).Address & ":" & Range("tAmount").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long")
            Range("uShortSizeInUSD").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tSizeInUSD").Offset(1, 0).Address & ":" & Range("tSizeInUSD").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long")
            Range("uShortGrossPL").Offset(i, 0).Value = Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tBaseCurrency").Offset(1, 0).Address & ":" & Range("tBaseCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Short") + Application.WorksheetFunction.SumIfs( _
                Range(Range("tGrossPL").Offset(1, 0).Address & ":" & Range("tGrossPL").Offset(zOpenTrades.Count, 0).Address), _
                Range(Range("tCounterCurrency").Offset(1, 0).Address & ":" & Range("tCounterCurrency").Offset(zOpenTrades.Count, 0).Address), _
                Range("uCurrency").Offset(i, 0).Value, _
                Range(Range("tBuySell").Offset(1, 0).Address & ":" & Range("tBuySell").Offset(zOpenTrades.Count, 0).Address), _
                "Long")
            With Range("uShortGrossPL").Offset(i, 0)
                If .Value < 0 Then
                    .Font.Color = zColorNegative
                ElseIf .Value = 0 Then
                    .Font.Color = zColorNeutral
                Else
                    .Font.Color = zColorPositive
                End If
            End With
            
            If Range("uLongAmount").Offset(i, 0).Value = Range("uShortAmount").Offset(i, 0).Value Then
                Range("uBuySell").Offset(i, 0).Value = "Flat"
                Range("uBuySell").Offset(i, 0).Font.Color = zColorNeutral
            ElseIf Range("uLongAmount").Offset(i, 0).Value > Range("uShortAmount").Offset(i, 0).Value Then
                Range("uBuySell").Offset(i, 0).Value = "Long"
                Range("uBuySell").Offset(i, 0).Font.Color = zColorPositive
            Else
                Range("uBuySell").Offset(i, 0).Value = "Short"
                Range("uBuySell").Offset(i, 0).Font.Color = zColorNegative
            End If
            Range("uNetAmount").Offset(i, 0).Value = Range("uLongAmount").Offset(i, 0).Value - Range("uShortAmount").Offset(i, 0).Value
            Range("uNetSizeInUSD").Offset(i, 0).Value = Range("uLongSizeInUSD").Offset(i, 0).Value - Range("uShortSizeInUSD").Offset(i, 0).Value
            Range("uNetGrossPL").Offset(i, 0).Value = Range("uLongGrossPL").Offset(i, 0).Value - Range("uShortGrossPL").Offset(i, 0).Value
            With Range("uNetGrossPL").Offset(i, 0)
                If .Value < 0 Then
                    .Font.Color = zColorNegative
                ElseIf .Value = 0 Then
                    .Font.Color = zColorNeutral
                Else
                    .Font.Color = zColorPositive
                End If
            End With
        Next i
    End If
End Sub

Private Sub ColorCurrencySummary()
    If zOpenTrades.Count > 0 Then
        
        ' Clears any existing formatting
        Range(Range("uAccountName").Offset(1, 0).Address & ":" & Range("uNetGrossPL").Offset _
            (zCurrencyCollection.Count, 0).Address).Interior.Color = xlNone
        Range(Range("uAccountName").Offset(1, 0).Address & ":" & Range("uNetGrossPL").Offset _
            (zCurrencyCollection.Count, 0).Address).Select
        
        ' Colors every other row a different color for improved readability
        For i = 1 To Selection.Rows.Count
            If i Mod 2 = 1 Then
                Selection.Rows(i).Interior.Color = zColorAlternatingRow
            End If
        Next
    End If
    Range("iStart").Select
End Sub

Private Sub ClearCurrencySummary()
    For Each zNamedCell In Names
        If InStr(1, zNamedCell.Name, "CurrencySummary_") > 0 Then Range(zNamedCell.Name).EntireRow.Delete
        If InStr(1, zNamedCell.RefersTo, "#REF!") > 0 Then zNamedCell.Delete
    Next zNamedCell
End Sub

Private Sub CheckCurrentState()
    If zTransport Is Nothing Then
        ' Starts the connection process
        Set zTransport = CreateObject("fxcore2.com.Transport")
        Set zSession = zTransport.createSession()
        
        ' Ensures that only a true status message get displayed
        If zLastSessionStatus <> 0 Then
            Range("iConnectionStatus").Value = zLastSessionStatus
        End If
    End If
End Sub

Private Sub zSession_SessionStatusChanged(ByVal xSessionStatus As SessionStatusCode)
    ' Displayes the current connection status under "Connection Status"
    zLastSessionStatus = xSessionStatus
    Range("iConnectionStatus").Value = GetStatusName(xSessionStatus)
    
    ' Colors the current connection status cell green if connected
    If GetStatusName(xSessionStatus) = "Connected" Then
        With Range("iConnectionStatus")
            .Font.Color = vbWhite
            .Interior.Color = zColorPositive
        End With
    Else
        With Range("iConnectionStatus")
            .Font.Color = vbWhite
            .Interior.Color = zColorNegative
        End With
    End If
End Sub

Private Sub zSession_LoginFailed(ByVal xErrorMessage As String)
    If xErrorMessage = "Wrong user name or password" Then
        Range("iMessage").Value = "Wrong Login or Password"
    ElseIf InStr(xErrorMessage, "Trading is closed for the weekend.") Then
        Range("iMessage").Value = "Trading is Closed for the Weekend"
    Else
        Range("iMessage").Value = xErrorMessage
    End If
    
    ' Allows the sub login to exit properly based on the connection status
    zReadyToStart = True
    zTablesLoaded = False
    zLoginSuccess = False
End Sub

Private Sub zSession_ManagerStatusChanged(ByVal xConnectionStatus As fxcore2_com.TableManagerStatus, ByVal xTableManagerObject As fxcore2_com.ITableManager)
    If xConnectionStatus = TableManagerStatus_TablesLoaded Then
        ' Allows the sub login to exit properly based on the connection status
        zReadyToStart = True
        zTablesLoaded = True
        zLoginSuccess = True
    ElseIf xConnectionStatus = TableManagerStatus_TablesLoadFailed Then
        ' Allows the sub login to exit properly based on the connection status
        zReadyToStart = True
        zTablesLoaded = False
        zLoginSuccess = True
    End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal xCellClicked As Range, Cancel As Boolean)
    On Error Resume Next
    ' Facilitates sorting based on a user double click of column headings
    If zLastSessionStatus = SessionStatusCode_Connected Then
        If xCellClicked.Value <> vbNullString Then
            ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Clear
            If xCellClicked.Row = Range("aAccountName").Row Then
                For Each zNamedCell In Names
                    If InStr(1, zNamedCell.Name, "Account_") > 0 Then zNamedCell.Delete
                Next zNamedCell
                For Each zCell In Range("aAccountName").Resize(, zLengthAccounts)
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
                Next zCell
                If zAccountsSortStatus = "Descending" Then
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zAccounts.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    zAccountsSortStatus = "Ascending"
                    xCellClicked.Value = "(" & ChrW(&H2191) & ") " & xCellClicked.Value
                Else
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zAccounts.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                        zAccountsSortStatus = "Descending"
                    xCellClicked.Value = "(" & ChrW(&H2193) & ") " & xCellClicked.Value
                End If
                With ActiveWorkbook.Worksheets("Account Dashboard").Sort
                    .SetRange Range(Range("aAccountName").Address & ":" & Range("aUsedMargin").Offset(zAccounts.Count, 0).Address)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Range("aAccountName").Resize(, zLengthAccounts).Interior.Color = vbBlack
                xCellClicked.Interior.Color = zColorNegative
                For i = 0 To zAccounts.Count - 1
                    Names.Add Name:="Account_" & Range("aAccountName").Offset(i + 1), RefersTo:=Range("aAccountName").Offset(i + 1)
                Next i
                Cancel = True
                Range("iStart").Select
                Call ColorAccounts
            ElseIf xCellClicked.Row = Range("tAccountName").Row Then
                For Each zNamedCell In Names
                    If InStr(1, zNamedCell.Name, "OpenTicket_") > 0 Then zNamedCell.Delete
                Next zNamedCell
                For Each zCell In Range("tAccountName").Resize(, zLengthOpenTrades)
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
                Next zCell
                If zOpenTradesSortStatus = "Descending" Then
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zOpenTrades.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    zOpenTradesSortStatus = "Ascending"
                    xCellClicked.Value = "(" & ChrW(&H2191) & ") " & xCellClicked.Value
                Else
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zOpenTrades.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                    zOpenTradesSortStatus = "Descending"
                    xCellClicked.Value = "(" & ChrW(&H2193) & ") " & xCellClicked.Value
                End If
                With ActiveWorkbook.Worksheets("Account Dashboard").Sort
                    .SetRange Range(Range("tTicketNumber").Address & ":" & Range("tUsedMargin").Offset(zOpenTrades.Count, 0).Address)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Range("tAccountName").Resize(, zLengthOpenTrades).Interior.Color = vbBlack
                xCellClicked.Interior.Color = zColorNegative
                For i = 0 To zOpenTrades.Count - 1
                    Names.Add Name:="OpenTicket_" & Range("tTicketNumber").Offset(i + 1), RefersTo:=Range("tTicketNumber").Offset(i + 1)
                Next i
                Cancel = True
                Range("iStart").Select
                Call ColorOpenTrades
            ElseIf xCellClicked.Row = Range("cAccountName").Row Then
                For Each zNamedCell In Names
                    If InStr(1, zNamedCell.Name, "ClosedTicket_") > 0 Then zNamedCell.Delete
                Next zNamedCell
                For Each zCell In Range("cAccountName").Resize(, zLengthClosedTrades)
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
                Next zCell
                If zClosedTradesSortStatus = "Descending" Then
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zClosedTrades.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    zClosedTradesSortStatus = "Ascending"
                    xCellClicked.Value = "(" & ChrW(&H2191) & ") " & xCellClicked.Value
                Else
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zClosedTrades.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                    zClosedTradesSortStatus = "Descending"
                    xCellClicked.Value = "(" & ChrW(&H2193) & ") " & xCellClicked.Value
                End If
                With ActiveWorkbook.Worksheets("Account Dashboard").Sort
                    .SetRange Range(Range("cTicketNumber").Address & ":" & Range("cRolloverInterest").Offset(zClosedTrades.Count, 0).Address)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Range("cAccountName").Resize(, zLengthClosedTrades).Interior.Color = vbBlack
                xCellClicked.Interior.Color = zColorNegative
                For i = 0 To zClosedTrades.Count - 1
                    Names.Add Name:="ClosedTicket_" & Range("cTicketNumber").Offset(i + 1), RefersTo:=Range("cTicketNumber").Offset(i + 1)
                Next i
                Cancel = True
                Range("iStart").Select
                Call ColorClosedTrades
            ElseIf xCellClicked.Row = Range("uAccountName").Row Then
                For Each zNamedCell In Names
                    If InStr(1, zNamedCell.Name, "CurrencyTicket_") > 0 Then zNamedCell.Delete
                Next zNamedCell
                For Each zCell In Range("uAccountName").Resize(, zLengthCurrency)
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2191) & ") ", "")
                    If zCell Like "*" & "(" & "*" Then zCell = Replace(zCell, "(" & ChrW(&H2193) & ") ", "")
                Next zCell
                If zCurrencySortStatus = "Descending" Then
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zCurrencyCollection.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    zCurrencySortStatus = "Ascending"
                    xCellClicked.Value = "(" & ChrW(&H2191) & ") " & xCellClicked.Value
                Else
                    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Add _
                        Key:=Range(xCellClicked.Offset(1, 0).Address & ":" & xCellClicked.Offset(zCurrencyCollection.Count, 0).Address), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                    zCurrencySortStatus = "Descending"
                    xCellClicked.Value = "(" & ChrW(&H2193) & ") " & xCellClicked.Value
                End If
                With ActiveWorkbook.Worksheets("Account Dashboard").Sort
                    .SetRange Range(Range("uCurrency").Address & ":" & Range("uNetGrossPL").Offset(zCurrencyCollection.Count, 0).Address)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                Range("uAccountName").Resize(, zLengthCurrency).Interior.Color = vbBlack
                xCellClicked.Interior.Color = zColorNegative
                For i = 1 To zCurrencyCollection.Count
                    Names.Add Name:="CurrencySummary_" & Range("uCurrency").Offset(i, 0).Value, RefersTo:=Range("uCurrencyTicket").Offset(i, 0)
                Next i
                Cancel = True
                Range("iStart").Select
                Call ColorCurrencySummary
            End If
        End If
    End If
End Sub

Private Sub SpeedUpVBA()
    ' Removes unnecessary features that create CPU load when the code is running
    ActiveWorkbook.Worksheets("Account Dashboard").DisplayPageBreaks = False
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
    End With
    With Application
        .ScreenUpdating = False
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .ErrorCheckingOptions.NumberAsText = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub ReturnToNormalVBA()
    ' Returns all features to the typical norm after code is done running
    ActiveWorkbook.Worksheets("Account Dashboard").Sort.SortFields.Clear
    With ActiveWindow
        .DisplayHeadings = True
    End With
    With Application
        .DisplayFullScreen = False
        .ScreenUpdating = True
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .ErrorCheckingOptions.NumberAsText = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Private Sub ErrorMessageText()
    ' Colors the "iMessage" cell red and selects it to alert the user
    With Range("iMessage")
        .Font.Color = vbWhite
        .Interior.Color = zColorNegative
        .Select
    End With
End Sub

Private Sub NoErrorMessageText()
    ' Colors the "iMessage" cell black and deleted any existing text
    With Range("iMessage")
        .Font.Color = vbWhite
        .Interior.Color = vbBlack
        .Value = vbNullString
    End With
End Sub

Private Sub Worksheet_Activate()
    Call CheckCurrentState
End Sub

Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Set zSession = Nothing
    Set zTransport = Nothing
    Call ReturnToNormalVBA
End Sub
