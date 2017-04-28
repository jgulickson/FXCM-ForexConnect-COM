Attribute VB_Name = "Other_SetGlobalVariables"
' All variable declarations
Public zAccountsSortStatus As String
Public zOpenTradesSortStatus As String
Public zClosedTradesSortStatus As String
Public zCurrencySortStatus As String

Public zUTCOffset As Integer
Public zLastRefreshTime As Date

Public zColorNegative As Long
Public zColorNeutral As Long
Public zColorPositive As Long
Public zColorAlternatingRow As Long
Public zColorTotalRow As Long
Public zColorTotalLine As Long

Public Sub SetGlobalVariables()
    
    ' Sorting variables
    zAccountsSortStatus = "Descending"
    zOpenTradesSortStatus = "Descending"
    zClosedTradesSortStatus = "Descending"
    zCurrencySortStatus = "Descending"
    
    ' Time variables
    zUTCOffset = (GetLocalToGMTDifference / 3600) * -1
    zLastRefreshTime = Now()

    ' Color variables
    zColorNegative = RGB(255, 0, 0)
    zColorNeutral = vbBlack
    zColorPositive = RGB(0, 153, 0)
    zColorAlternatingRow = RGB(220, 230, 241)
    zColorTotalRow = xlNone
    zColorTotalLine = vbBlack
    
End Sub
