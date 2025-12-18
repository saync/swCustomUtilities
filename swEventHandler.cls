VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "swEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents app As SldWorks.SldWorks
Attribute app.VB_VarHelpID = -1
Private WithEvents part As SldWorks.partDoc
Attribute part.VB_VarHelpID = -1
Private WithEvents assembly As SldWorks.AssemblyDoc
Attribute assembly.VB_VarHelpID = -1
Private WithEvents drawing As SldWorks.DrawingDoc
Attribute drawing.VB_VarHelpID = -1

Dim doc As SldWorks.ModelDoc2
Dim docExt As SldWorks.ModelDocExtension


Private Sub Class_Initialize()
    logger.add "swEventHandler Initialized..."
    Set app = swApp
    Set doc = swApp.ActivateDoc
    
    updateModels
    
End Sub


Private Sub updateModels()
    Select Case doc.GetType
        Case swDocNONE
            Set part = Nothing
            Set assembly = Nothing
            Set drawing = Nothing
        Case swDocPART
            Set part = doc
            Set assembly = Nothing
            Set drawing = Nothing
        Case swDocASSEMBLY
            Set part = Nothing
            Set assembly = doc
            Set drawing = Nothing
        Case swDocDRAWING
            Set part = Nothing
            Set assembly = Nothing
            Set drawing = doc
    End Select
    
    Set docExt = doc.Extension
    
End Sub


Private Function app_ActiveModelDocChangeNotify() As Long
    Set doc = app.ActivateDoc
    updateModels
    
    logger.add "Active model set to: " & doc.GetTitle
    
End Function


Private Function app_FileOpenPostNotify(ByVal FileName As String) As Long
    logger.add "File opened: " & FileName

End Function


Private Function part_FileSavePostNotify(ByVal saveType As Long, ByVal FileName As String) As Long
    Select Case saveType
        Case swFileSave
            logger.add "file save"
        Case swFileSaveAs
            ' File save as routines
        Case swFileSaveAsCopy
            ' File save as copy routines
        Case swFileSaveAsCopyAndOpen
            ' File save as copy and open routines
    End Select
    
End Function

