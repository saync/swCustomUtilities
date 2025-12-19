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

Dim utilities As swCustomUtilities.UtilitiesOptions
Dim doc As SldWorks.ModelDoc2


Public Sub initialize(ByRef swApp As SldWorks.SldWorks, ByRef utilityOptions As swCustomUtilities.UtilitiesOptions)
    logger.add "swEventHandler Initialized..."
    Set app = swApp
    Set doc = swApp.ActiveDoc
    Set utilities = utilityOptions
    
    updateModels
    
End Sub


Private Sub updateModels()
    If Not doc Is Nothing Then
        Select Case (doc.GetType)
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
    End If
    
End Sub


Private Function app_ActiveModelDocChangeNotify() As Long
    Set doc = app.ActiveDoc
    updateModels
    
    If doc Is Nothing Then
        logger.add "Active model set to: Nothing"
    Else
        logger.add "Active model set to: " & doc.GetTitle
    End If
    
End Function


Private Function app_FileOpenNotify2(ByVal fileName As String) As Long
    logger.add "File opened: " & fileName
    
    logger.add module_SaveExport.fileSummary(doc)
    
End Function


Private Function part_FileSavePostNotify(ByVal saveType As Long, ByVal fileName As String) As Long
    Select Case saveType
        Case swFileSave
            logger.add "File saved: " & fileName
            
            If utilities.saveExport Then
                If UCase((module_SaveExport.getFileExt(fileName))) = "SLDPRT" Then
                    module_SaveExport.exportStep doc
                    
                End If
            End If
            
        Case swFileSaveAs
            ' File save as routines
        Case swFileSaveAsCopy
            ' File save as copy routines
        Case swFileSaveAsCopyAndOpen
            ' File save as copy and open routines
    End Select
    
End Function
