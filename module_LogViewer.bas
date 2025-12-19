Attribute VB_Name = "module_LogViewer"
Option Explicit

Public logger As swCustomUtilities.DebugLogger


Public Sub startup(ByRef app As SldWorks.SldWorks)
    Dim startupMsg As String
    Dim baseVer As String
    Dim currentVer As String

    Set logger = New swCustomUtilities.DebugLogger
    startupMsg = ""

    If app.ApplicationType = swApplicationType_3DEXPERIENCE Then
        startupMsg = "Solidworks 3D Experience application" & vbCrLf
    ElseIf app.ApplicationType = swApplicationType_Desktop Then
        startupMsg = "Solidworks Desktop application" & vbCrLf
    End If
    
    app.GetBuildNumbers2 baseVer, currentVer, Empty
    
    startupMsg = startupMsg & "Version: " & baseVer & "." & currentVer & vbCrLf & _
    "- " & app.GetExecutablePath & " [PID] " & app.GetProcessID & vbCrLf & _
    String(94, "/")

    logger.add startupMsg

    startupMsg = ""
End Sub


Public Function multiline(ByVal strArry As Variant, Optional ByVal level As Integer = 1) As String
    Dim returnstr As String
    Dim i, j As Integer
    
    returnstr = ""
    
    For i = LBound(strArry) To UBound(strArry)
        For j = 1 To level
            returnstr = returnstr & vbTab
        Next j
        
        returnstr = returnstr & strArry(i) & vbCrLf
    Next i
    
    multiline = returnstr
    
End Function



