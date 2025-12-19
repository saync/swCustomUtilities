# swCustomUtilities
Custom solidworks vba utilitly macros

## Live Logging
<img width="1240" height="960" alt="image" src="https://github.com/user-attachments/assets/0af0d949-f1d3-4d8d-ac45-0a846774a812" />

Logging form used to debug and display information in a live environment. Add a public instance of the DebugLogger class and use the _add_ method to log messages to both the vba environment immediate window as well as the live logger.

```
Public Sub add(ByVal message As String, Optional ByVal severity As String = "INFO")
    Dim logMessage As String
    Dim messageSeverity As String
    
    Select Case UCase(severity)
        Case "0", "I", "INFO"
            messageSeverity = "INFO"
        Case "1", "W", "WARNING"
            messageSeverity = "WARNING"
        Case "2", "E", "ERROR"
            messageSeverity = "ERROR"
        Case Else
            messageSeverity = "NONE"
    End Select

    logMessage = DateTime.Now & " [" & messageSeverity & "]" & " >> " & message
    Me.Log = Me.Log & logMessage & vbCrLf
    updateLog
    Debug.Print logMessage
    
End Sub
```

## Adding Macro Button
<img width="322" height="330" alt="image" src="https://github.com/user-attachments/assets/6be95794-2d9f-4aaa-bae3-764c14aec7c5" />

Tools > Customize > Drag and drop "New Macro Button" onto toolbar.
