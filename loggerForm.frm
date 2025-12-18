VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loggerForm 
   Caption         =   "Log"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "loggerForm.frx":0000
End
Attribute VB_Name = "LoggerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub UserForm_Initialize()
    Me.Width = 400
    Me.Height = 600
    
End Sub

Private Sub UserForm_Resize()
    logText.Width = Me.Width
    logText.Height = Me.Height

End Sub
