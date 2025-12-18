VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugLogger"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public logForm As LoggerForm
Public Log As String


Private Sub Class_Initialize()
    Set Me.logForm = New LoggerForm
    
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


Public Sub add(ByVal message As String)
    Dim logMessage As String
    
    logMessage = Now & " >> " & message
    Me.Log = Me.Log & logMessage & vbCrLf
    updateLog
    Debug.Print logMessage
    
End Sub

