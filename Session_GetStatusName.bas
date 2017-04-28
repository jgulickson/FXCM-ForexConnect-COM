Attribute VB_Name = "Session_GetStatusName"
Public Function GetStatusName(xConnectionStatus As SessionStatusCode)
    Select Case xConnectionStatus
        Case SessionStatusCode_Connected
            GetStatusName = "Connected"
        Case SessionStatusCode_Disconnected
            GetStatusName = "Disconnected"
        Case SessionStatusCode_Connecting
            GetStatusName = "Connecting..."
        Case SessionStatusCode_TradingSessionRequested
            GetStatusName = "Trading Session Requested..."
        Case SessionStatusCode_Disconnecting
            GetStatusName = "Disconnecting..."
        Case SessionStatusCode_SessionLost
            GetStatusName = "Session Lost"
        Case SessionStatusCode_PriceSessionReconnecting
            GetStatusName = "Price Session Reconnecting..."
        Case SessionStatusCode_Unknown
            GetStatusName = "Unknown"
        Case Else
            GetStatusName = CStr(xConnectionStatus)
    End Select
End Function
