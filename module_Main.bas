Attribute VB_Name = "module_Main"
Option Explicit

Private swApp As SldWorks.SldWorks
Private utilities As New UtilitiesOptions
Dim eventHandler As swCustomUtilities.swEventHandler


Sub main()
    utilities.saveExport = True
    
    Set swApp = Application.SldWorks
    
    module_LogViewer.startup swApp
    
    If utilities.saveExport Then
        module_SaveExport.init swApp
    End If
        
    Set eventHandler = New swCustomUtilities.swEventHandler
    eventHandler.initialize swApp, utilities
    
End Sub
