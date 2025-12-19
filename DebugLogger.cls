VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public logForm As loggerForm
Public Log As String


Private Sub Class_Initialize()
    Set Me.logForm = New loggerForm
    
    Me.Log = ""
    Me.add "Logger initialized..."
    
    showLog
    
End Sub


Public Sub showLog()
    Me.logForm.Show vbModeless
    
End Sub


Public Sub hideLog()
    Me.logForm.Hide
    
End Sub


Private Sub updateLog()
    Me.logForm.logText.Value = Me.Log
    
End Sub


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

