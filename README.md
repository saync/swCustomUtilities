# swCustomUtilities
Custom solidworks vba utilitly macros.

## Live Logging
<img width="1240" height="960" alt="image" src="https://github.com/user-attachments/assets/0af0d949-f1d3-4d8d-ac45-0a846774a812" />

Logging form used to debug and display information in a live environment. A public instance of the DebugLogger class that uses the _add_ method to log messages to both the vba environment immediate window as well as the live logger. Severity can be set for custom error handleing.

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

## Auto STEP file export

Uses the part _FileSavePostNotify_ event to automatically export step files to the same directory of the part file.


## Adding Macro Button
<img width="322" height="330" alt="image" src="https://github.com/user-attachments/assets/6be95794-2d9f-4aaa-bae3-764c14aec7c5" />

Tools > Customize. Drag and drop "New Macro Button" onto the toolbar and direct it to the macro file.
