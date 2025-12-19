Attribute VB_Name = "module_SaveExport"
Option Explicit

Dim fso As Scripting.FileSystemObject
Private app As SldWorks.SldWorks

Public Sub init(ByRef swApp As SldWorks.SldWorks)
    Set fso = New Scripting.FileSystemObject
    Set app = swApp

End Sub


Public Function getFileExt(ByVal fileName As String) As String
    getFileExt = fso.GetExtensionName(fileName)
 
End Function


Private Function replaceExt(ByVal fileName As String, ByVal newExt As String) As String
    Dim revStr As String
    Dim fileExt As String
    
    fileExt = StrReverse(fso.GetExtensionName(fileName))
    revStr = StrReverse(fileName)
    newExt = StrReverse(newExt)
    
    replaceExt = StrReverse((Replace(revStr, fileExt, newExt, 1, 1)))
    
End Function


Public Sub exportStep(ByRef modelDoc As ModelDoc2, Optional ByVal outputFile As String = "")
    Dim filePath As String
    Dim fileName As String
    Dim saveErrors As Long
    Dim saveWarnings As Long
    Dim advancedSaveOptions As AdvancedSaveAsOptions
    Dim saveSuccess As Boolean
    
    If outputFile = "" Then
        filePath = modelDoc.GetPathName
    Else
        filePath = outputFile
    End If
    
    fileName = fso.GetFileName(filePath)
    filePath = fso.GetParentFolderName(filePath)
    
    If Not fso.FolderExists(filePath) Then
        fso.CreateFolder (filePath)
        logger.add "Folderpath created: " & filePath, "WARNING"
        
    End If
    
    fileName = replaceExt(fileName, "STEP")
    filePath = fso.BuildPath(filePath, fileName)
    
    Set advancedSaveOptions = modelDoc.Extension.GetAdvancedSaveAsOptions(0)
    
    saveSuccess = modelDoc.Extension.SaveAs3(filePath, _
        swSaveAsCurrentVersion, _
        swSaveAsOptions_Silent, _
        Nothing, _
        advancedSaveOptions, _
        saveErrors, _
        saveWarnings _
        )
        
    If saveSuccess Then
        logger.add "File exported successful: " & filePath
        logger.add "File save warnings: " & saveWarnings, "WARNING"
    Else
        logger.add "File exported unsuccessful: " & filePath
        logger.add "File save warnings: " & saveWarnings, "WARNING"
        logger.add "File save errors: " & saveErrors, "ERROR"
    End If
    
End Sub


Public Function fileSummary(ByRef doc As ModelDoc2) As String
    Dim mass As MassProperty2
    Dim swMassUnits As Integer
    Dim massUnits As String
    Dim lengthUnits As String
    Dim docType As String
    Dim summary() As Variant
    Dim summary1() As Variant
    Dim summary2() As Variant
    Dim summary3() As Variant
    Dim cX, cY, cZ As Double
    Dim Ixx, Ixy, Ixz As Double
    Dim Iyx, Iyy, Iyz As Double
    Dim Izx, Izy, Izz As Double
    Dim Px, Py, Pz As Double
    
    fileSummary = "[      File Summary      ]" & vbCrLf
    
    Set mass = doc.Extension.CreateMassProperty2
    
    Select Case doc.GetType
        Case swDocPART
            docType = "Part"
        Case swDocASSEMBLY
            docType = "Assembly"
        Case swDocDRAWING
            docType = "Drawing"
        Case Else
            docType = ""
    End Select
    
    Select Case (doc.GetUnits(0))
        Case swANGSTROM
            lengthUnits = "Å"
        Case swCM
            lengthUnits = "cm"
        Case swFEET
            lengthUnits = "ft"
        Case swCM
            lengthUnits = "cm"
        Case swINCHES
            lengthUnits = "in"
        Case swMETER
            lengthUnits = "m"
        Case swMICRON
            lengthUnits = "µm"
        Case swMIL
            lengthUnits = "mil"
        Case swMM
            lengthUnits = "mm"
        Case swNANOMETER
            lengthUnits = "nm"
        Case swUIN
            lengthUnits = "µin"
        Case Else
            lengthUnits = ""
    End Select
    
    
    summary1 = Array( _
        String(60, "V"), _
        "Title:     " & doc.SummaryInfo(swSumInfoTitle), _
        "Subject:   " & doc.SummaryInfo(swSumInfoSubject), _
        "Author:    " & doc.SummaryInfo(swSumInfoAuthor), _
        "Created:   " & doc.SummaryInfo(swSumInfoCreateDate), _
        "Saved By:  " & doc.SummaryInfo(swSumInfoSavedBy), _
        "Saved:     " & doc.SummaryInfo(swSumInfoSaveDate), _
        "Keywords:  " & doc.SummaryInfo(swSumInfoKeywords), _
        "Comments:  " & doc.SummaryInfo(swSumInfoComment), _
        String(60, "-"), _
        "Doc Type:           " & docType _
        )
        
    If Not mass Is Nothing Then
        mass.UseSystemUnits = False
        swMassUnits = doc.Extension.GetUserPreferenceInteger(swUnitsMassPropMass, swDetailingNoOptionSpecified)
        
        Select Case swMassUnits
            Case swUnitsMassPropMass_Grams
                massUnits = "g"
            Case swUnitsMassPropMass_Kilograms
                massUnits = "kg"
            Case swUnitsMassPropMass_Milligrams
                massUnits = "mg"
            Case swUnitsMassPropMass_Pounds
                massUnits = "lbs"
            Case Else
                massUnits = ""
        End Select
        
        cX = mass.CenterOfMass(0)
        cY = mass.CenterOfMass(1)
        cZ = mass.CenterOfMass(2)
        
        Px = mass.PrincipalMomentsOfInertia(0)
        Py = mass.PrincipalMomentsOfInertia(1)
        Pz = mass.PrincipalMomentsOfInertia(2)
        
        Ixx = mass.PrincipalAxesOfInertia(0)(0)
        Ixy = mass.PrincipalAxesOfInertia(0)(1)
        Ixz = mass.PrincipalAxesOfInertia(0)(2)
        Iyx = mass.PrincipalAxesOfInertia(1)(0)
        Iyy = mass.PrincipalAxesOfInertia(1)(1)
        Iyz = mass.PrincipalAxesOfInertia(1)(2)
        Izx = mass.PrincipalAxesOfInertia(2)(0)
        Izy = mass.PrincipalAxesOfInertia(2)(1)
        Izz = mass.PrincipalAxesOfInertia(2)(2)
        
        summary2 = Array( _
            "Mass:               " & mass.mass & " " & massUnits, _
            "Volume:             " & mass.Volume & " " & lengthUnits & "^3", _
            "Surface Area:       " & Round(mass.SurfaceArea, 4) & " " & lengthUnits & "^2", _
            "Center of Mass:     [ " & pad(cX) & ", " & pad(cY) & ", " & pad(cZ) & " ] (x,y,z)", _
            "Moments of Inertia: [ " & pad(Px) & ", " & pad(Py) & ", " & pad(Pz) & " ] (x,y,z)", _
            "Axes of Inertia:    " _
            )
            
        summary3 = Array( _
            "[ " & pad(Ixx) & ", " & pad(Ixy) & ", " & pad(Ixz) & " ]", _
            "[ " & pad(Iyx) & ", " & pad(Iyy) & ", " & pad(Iyz) & " ]", _
            "[ " & pad(Izx) & ", " & pad(Izy) & ", " & pad(Izz) & " ]", _
            String(20, "."), _
            "[ Ixx, Ixy, Ixz ]", _
            "[ Iyx, Iyy, Iyz ]", _
            "[ Izx, Izy, Izz ]" _
            )
    
    End If
    
    fileSummary = fileSummary & module_LogViewer.multiline(summary1, 1)
    
    If Not IsEmpty(summary2) Then
        fileSummary = fileSummary & module_LogViewer.multiline(summary2, 1) & module_LogViewer.multiline(summary3, 2)
    End If
    
    fileSummary = fileSummary & vbCrLf & vbTab & String(60, "-")
End Function


Private Function pad(ByVal num As Variant) As String
    pad = Format(num, "0.0000E+00")

    If num >= 0 Then
        pad = " " & pad
    End If
End Function
    
