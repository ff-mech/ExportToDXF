' =============================================================================
' ExportToDXF_rE — Revision E (Images Removed ONLY)
' =============================================================================
' Based on your original Revision E.
' Removed only:
'   - Check image saving
'   - Failed check image saving
'   - Check / FAILED folder creation
'   - Image-related logging and messaging
' Everything else is preserved.
' =============================================================================

Option Explicit

Dim swApp As SldWorks.SldWorks

Private Type PartAnalysis
    HasSheetMetal       As Boolean
    HasBends            As Boolean
    HasRolledBack       As Boolean
    flatPatternName     As String
    MultipleConfigs     As Boolean
End Type

' Per-file result returned by ProcessPart
Private Enum PartResult
    prExported = 0
    prFailed = 1
    prSkipped = 2
End Enum

' --- ENTRY POINT -------------------------------------------------------------

Sub main()

    Set swApp = Application.SldWorks

    ' -- Phase 1: User Input --------------------------------------------------

    Dim sourceFolder As String
    sourceFolder = InputBox("Paste folder path containing .sldprt files:", _
                            "DXF Export - Source Folder")

    If Trim$(sourceFolder) = "" Then
        swApp.SendMsgToUser2 "No source folder provided. Exiting.", _
                             swMbWarning, swMbOk
        Exit Sub
    End If

    sourceFolder = NormalizeFolder(sourceFolder)

    If Not FolderExists(sourceFolder) Then
        swApp.SendMsgToUser2 "Source folder does not exist:" & vbCrLf & _
                             sourceFolder, swMbStop, swMbOk
        Exit Sub
    End If

    Dim destFolder As String
    destFolder = InputBox("Paste destination folder for DXF export:", _
                          "DXF Export - Destination Folder", sourceFolder)

    If Trim$(destFolder) = "" Then
        swApp.SendMsgToUser2 "No destination folder provided. Exiting.", _
                             swMbWarning, swMbOk
        Exit Sub
    End If

    destFolder = NormalizeFolder(destFolder)

    If Not FolderExists(destFolder) Then
        swApp.SendMsgToUser2 "Destination folder does not exist:" & vbCrLf & _
                             destFolder, swMbStop, swMbOk
        Exit Sub
    End If

    Dim logPath As String
    logPath = destFolder & "DXF_Export_Log.txt"

    Dim overallStart As Single
    overallStart = Timer

    ' -- Phase 3: Process Files -----------------------------------------------

    Dim fileName As String
    fileName = Dir(sourceFolder & "*.sldprt")

    Dim exportCount As Long
    Dim failCount As Long
    Dim skipCount As Long
    Dim processedCount As Long

    Dim failLog As String
    Dim warnLog As String
    Dim okLog As String

    failLog = ""
    warnLog = ""
    okLog = ""

    Do While fileName <> ""

        processedCount = processedCount + 1

        Dim partLog As String
        Dim result As PartResult

        result = ProcessPart(fileName, sourceFolder, destFolder, partLog)

        Select Case result
            Case prExported
                okLog = okLog & partLog & vbCrLf
                exportCount = exportCount + 1
            Case prFailed
                failLog = failLog & partLog & vbCrLf
                failCount = failCount + 1
            Case prSkipped
                warnLog = warnLog & partLog & vbCrLf
                skipCount = skipCount + 1
        End Select

        fileName = Dir
    Loop

    ' -- Phase 4: Reporting ---------------------------------------------------

    Dim overallElapsed As Double
    overallElapsed = ElapsedSeconds(overallStart)

    Dim avgPerPart As Double
    If processedCount > 0 Then
        avgPerPart = overallElapsed / processedCount
    Else
        avgPerPart = 0#
    End If

    Dim logFile As Integer
    logFile = FreeFile
    Open logPath For Output As #logFile

    Print #logFile, "DXF EXPORT LOG"
    Print #logFile, "Started: " & Now
    Print #logFile, "Source Folder: " & sourceFolder
    Print #logFile, "Destination Folder: " & destFolder
    Print #logFile, String(70, "=")
    Print #logFile, ""

    Print #logFile, "ERRORS / FAILURES FIRST"
    Print #logFile, String(70, "-")
    If failLog <> "" Then
        Print #logFile, failLog
    Else
        Print #logFile, "None"
        Print #logFile, ""
    End If

    Print #logFile, "SKIPS / WARNINGS"
    Print #logFile, String(70, "-")
    If warnLog <> "" Then
        Print #logFile, warnLog
    Else
        Print #logFile, "None"
        Print #logFile, ""
    End If

    Print #logFile, "SUCCESSFUL EXPORTS"
    Print #logFile, String(70, "-")
    If okLog <> "" Then
        Print #logFile, okLog
    Else
        Print #logFile, "None"
        Print #logFile, ""
    End If

    Print #logFile, "SUMMARY"
    Print #logFile, String(70, "-")
    Print #logFile, "Processed:            " & processedCount
    Print #logFile, "Exported:             " & exportCount
    Print #logFile, "Failed:               " & failCount
    Print #logFile, "Skipped:              " & skipCount
    Print #logFile, "Average Time / Part:  " & FormatSeconds(avgPerPart)
    Print #logFile, "Total Time:           " & FormatSeconds(overallElapsed)
    Print #logFile, String(70, "=")

    Close #logFile

    swApp.SendMsgToUser2 _
        "Done!" & vbCrLf & vbCrLf & _
        "Processed: " & processedCount & vbCrLf & _
        "Exported: " & exportCount & vbCrLf & _
        "Failed: " & failCount & vbCrLf & _
        "Skipped: " & skipCount & vbCrLf & _
        "Average Time / Part: " & FormatSeconds(avgPerPart) & vbCrLf & _
        "Total Time: " & FormatSeconds(overallElapsed) & vbCrLf & vbCrLf & _
        "Log saved to:" & vbCrLf & logPath, _
        swMbInformation, swMbOk

End Sub

' --- PER-FILE PROCESSOR ------------------------------------------------------

Private Function ProcessPart(ByVal fileName As String, _
                             ByVal sourceFolder As String, _
                             ByVal destFolder As String, _
                             ByRef partLog As String) As PartResult

    Dim partStart As Single
    partStart = Timer

    Dim filePath As String
    filePath = sourceFolder & fileName

    Dim fileBaseName As String
    fileBaseName = Left$(fileName, InStrRev(fileName, ".") - 1)

    Dim dxfPath As String
    dxfPath = destFolder & fileBaseName & ".dxf"

    Dim preOpenDocs As Object
    Set preOpenDocs = GetOpenDocDict()

    partLog = "FILE: " & fileName & vbCrLf

    Dim errors As Long
    Dim warnings As Long
    Dim swModel As SldWorks.ModelDoc2
    Dim swPart As SldWorks.PartDoc

    Set swModel = swApp.OpenDoc6(filePath, swDocPART, _
                                 swOpenDocOptions_Silent, "", errors, warnings)

    If swModel Is Nothing Then
        partLog = partLog & "  FAIL - Could not open file (OpenDoc6 error: " & _
                  CStr(errors) & ")" & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
        GoTo Cleanup
    End If

    Set swPart = swModel

    ' -- Step 1: Single Rebuild -----------------------------------------------
    If Not FastRebuild(swModel, partLog, True) Then
        partLog = partLog & "  FAIL - Initial rebuild failed." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
        GoTo Cleanup
    End If

    ' -- Step 2: Analyze Part -------------------------------------------------
    Dim info As PartAnalysis
    info = AnalyzePart(swModel)

    If Not info.HasSheetMetal Then
        partLog = partLog & "  SKIP - Not a sheet metal part." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prSkipped
        GoTo Cleanup
    End If

    If info.MultipleConfigs Then
        partLog = partLog & "  WARNING - Multiple configurations detected." & vbCrLf
    End If

    If info.flatPatternName = "" Then
        partLog = partLog & "  FAIL - No Flat-Pattern feature found." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
        GoTo Cleanup
    End If

    ' -- Step 3: Rollback Correction ------------------------------------------
    If info.HasRolledBack Then
        partLog = partLog & "  WARN - Rollback state detected. Attempting to roll forward." & vbCrLf

        On Error Resume Next
        swModel.FeatureManager.EditRollback swMoveRollbackBarToEnd, ""
        FastRebuild swModel, partLog, True
        On Error GoTo 0

        info = AnalyzePart(swModel)
        partLog = partLog & "  INFO - Rolled forward, continuing." & vbCrLf
    Else
        partLog = partLog & "  INFO - Rollback bar already at end." & vbCrLf
    End If

    ' -- Step 4: Flatten + Validate + Export ----------------------------------
    If info.HasBends Then

        If Not FlattenPart(swModel, info.flatPatternName, partLog) Then
            partLog = partLog & "  FAIL - Could not flatten part." & vbCrLf
            AppendElapsed partLog, partStart
            ProcessPart = prFailed
            GoTo Cleanup
        End If

        If Not FastRebuild(swModel, partLog, True) Then
            partLog = partLog & "  FAIL - Rebuild failed after flatten." & vbCrLf
            AppendElapsed partLog, partStart
            ProcessPart = prFailed
            GoTo Cleanup
        End If

        Dim flatErrText As String
        flatErrText = ""

        If CheckFeatureErrors(swModel, flatErrText, info.flatPatternName) Then
            partLog = partLog & "  FAIL - Feature tree error(s) after flatten:" & _
                      vbCrLf & flatErrText
            AppendElapsed partLog, partStart
            ProcessPart = prFailed
            GoTo Cleanup
        End If

        partLog = partLog & "  INFO - Flat pattern validation passed." & vbCrLf

    Else
        partLog = partLog & "  INFO - No bends detected, skipping flat pattern validation." & vbCrLf
    End If

    ' -- Step 10: Export ------------------------------------------------------
    Dim alignArr(11) As Double
    alignArr(0) = 0#:  alignArr(1) = 0#:  alignArr(2) = 0#
    alignArr(3) = 1#:  alignArr(4) = 0#:  alignArr(5) = 0#
    alignArr(6) = 0#:  alignArr(7) = 1#:  alignArr(8) = 0#
    alignArr(9) = 0#:  alignArr(10) = 0#: alignArr(11) = 1#

    Dim smOptions As Long
    smOptions = 71

    Dim bRet As Boolean
    bRet = swPart.ExportToDWG2(dxfPath, filePath, 1, True, alignArr, _
                               False, False, smOptions, Empty)

    If bRet Then
        partLog = partLog & "  OK - Exported to: " & dxfPath & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prExported
    Else
        partLog = partLog & "  FAIL - ExportToDWG2 returned False." & vbCrLf
        AppendElapsed partLog, partStart
        ProcessPart = prFailed
    End If

Cleanup:
    On Error Resume Next

    If Not swModel Is Nothing Then
        swApp.CloseDoc swModel.GetTitle
        Set swModel = Nothing
        Set swPart = Nothing
    End If

    CloseExtraDocs preOpenDocs
    Set preOpenDocs = Nothing

    On Error GoTo 0

End Function

' --- ANALYZE PART ------------------------------------------------------------

Private Function AnalyzePart(ByVal swModel As SldWorks.ModelDoc2) As PartAnalysis

    Dim result As PartAnalysis
    Dim swFeat As SldWorks.Feature

    Set swFeat = swModel.FirstFeature

    Do While Not swFeat Is Nothing

        Dim featType As String
        featType = LCase$(swFeat.GetTypeName2)

        Select Case featType
            Case "sheetmetal"
                result.HasSheetMetal = True

            Case "flatpattern"
                If result.flatPatternName = "" Then
                    result.flatPatternName = swFeat.Name
                End If

            Case "edgeflange", "sketchbend", "foldfeature", _
                 "jog", "loftbend", "baseflangewall", _
                 "hem", "miterflange", "crossbreak", _
                 "smbaseflangewall"
                result.HasBends = True
        End Select

        If swFeat.IsRolledBack Then
            result.HasRolledBack = True
        End If

        Set swFeat = swFeat.GetNextFeature
    Loop

    On Error Resume Next

    Dim vConfNames As Variant
    vConfNames = swModel.GetConfigurationNames

    If Not IsEmpty(vConfNames) Then
        Dim userCount As Long
        Dim i As Long
        userCount = 0

        For i = LBound(vConfNames) To UBound(vConfNames)
            Dim upperName As String
            upperName = UCase$(Trim$(CStr(vConfNames(i))))

            If InStr(upperName, "SM-FLAT-PATTERN") = 0 And _
               InStr(upperName, "FLAT-PATTERN") = 0 Then
                userCount = userCount + 1
            End If
        Next i

        result.MultipleConfigs = (userCount > 1)
    End If

    On Error GoTo 0

    AnalyzePart = result

End Function

' --- CHECK FLAT PATTERN ERROR ------------------------------------------------

Private Function CheckFeatureErrors(ByVal swModel As SldWorks.ModelDoc2, _
                                    ByRef errorText As String, _
                                    ByVal flatPatternName As String) As Boolean

    CheckFeatureErrors = False
    errorText = ""

    On Error Resume Next

    Dim flatFeat As SldWorks.Feature
    Set flatFeat = swModel.FeatureByName(flatPatternName)

    If flatFeat Is Nothing Then
        errorText = "    - Flat-Pattern feature not found: " & flatPatternName & vbCrLf
        CheckFeatureErrors = True
        Exit Function
    End If

    Dim errCode As Long
    Dim isWarning As Boolean

    errCode = flatFeat.GetErrorCode2(isWarning)

    If Err.Number = 0 Then
        If errCode <> 0 And Not isWarning Then
            CheckFeatureErrors = True
            errorText = "    - Feature: " & flatFeat.Name & _
                " | Type: " & flatFeat.GetTypeName2 & _
                " | ErrorCode: " & CStr(errCode) & _
                " | IsWarning: " & CStr(isWarning) & vbCrLf
        End If
    End If

    Err.Clear
    On Error GoTo 0

End Function

' --- FLATTEN -----------------------------------------------------------------

Private Function FlattenPart(ByVal swModel As SldWorks.ModelDoc2, _
                             ByVal flatPatternName As String, _
                             ByRef partLog As String) As Boolean

    On Error GoTo EH

    Dim flatFeat As SldWorks.Feature
    Set flatFeat = swModel.FeatureByName(flatPatternName)

    If flatFeat Is Nothing Then
        partLog = partLog & "  ERROR - Flat-Pattern feature not found: " & _
                  flatPatternName & vbCrLf
        FlattenPart = False
        Exit Function
    End If

    partLog = partLog & "  INFO - Flattening: " & flatFeat.Name & vbCrLf

    swModel.ClearSelection2 True
    If flatFeat.Select2(False, 0) Then
        swModel.EditUnsuppress2
        partLog = partLog & "  INFO - Flattened via EditUnsuppress2." & vbCrLf
        FlattenPart = True
        Exit Function
    End If

    partLog = partLog & "  INFO - Select2 failed, using SetSuppression2 fallback." & vbCrLf

    Dim confNames(0) As String
    confNames(0) = swModel.ConfigurationManager.ActiveConfiguration.Name

    If flatFeat.SetSuppression2(swUnSuppressFeature, swSpecifyConfiguration, _
                                confNames) = False Then
        partLog = partLog & "  ERROR - SetSuppression2 fallback also failed." & vbCrLf
        FlattenPart = False
        Exit Function
    End If

    FlattenPart = True
    Exit Function

EH:
    partLog = partLog & "  ERROR - Exception during flatten: " & _
              Err.Description & vbCrLf
    FlattenPart = False

End Function

' --- FAST REBUILD ------------------------------------------------------------

Private Function FastRebuild(ByVal swModel As SldWorks.ModelDoc2, _
                             ByRef partLog As String, _
                             Optional ByVal topLevelOnly As Boolean = True) As Boolean

    On Error GoTo EH

    Dim ok As Boolean
    ok = swModel.ForceRebuild3(topLevelOnly)

    If ok Then
        partLog = partLog & "  INFO - Rebuild succeeded." & vbCrLf
    Else
        partLog = partLog & "  WARN - ForceRebuild3 returned False." & vbCrLf
    End If

    FastRebuild = True
    Exit Function

EH:
    partLog = partLog & "  ERROR - Exception during rebuild: " & _
              Err.Description & vbCrLf
    FastRebuild = False

End Function

Private Function SilentRebuild(ByVal swModel As SldWorks.ModelDoc2, _
                               Optional ByVal topLevelOnly As Boolean = True) As Boolean

    On Error GoTo EH

    swModel.ForceRebuild3 topLevelOnly
    SilentRebuild = True
    Exit Function

EH:
    SilentRebuild = False

End Function

' --- OPEN DOC TRACKING -------------------------------------------------------

Private Function GetOpenDocDict() As Object

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim vDocs As Variant
    vDocs = swApp.GetDocuments

    If IsEmpty(vDocs) Then
        Set GetOpenDocDict = dict
        Exit Function
    End If

    Dim i As Long
    For i = 0 To UBound(vDocs)
        Dim swDoc As SldWorks.ModelDoc2
        Set swDoc = vDocs(i)

        If Not swDoc Is Nothing Then
            Dim docPath As String
            docPath = LCase$(swDoc.GetPathName)
            If docPath <> "" And Not dict.Exists(docPath) Then
                dict.Add docPath, True
            End If
        End If
    Next i

    Set GetOpenDocDict = dict

End Function

Private Sub CloseExtraDocs(ByVal baselineDict As Object)

    On Error Resume Next

    If baselineDict Is Nothing Then Exit Sub

    Dim vDocs As Variant
    vDocs = swApp.GetDocuments

    If IsEmpty(vDocs) Then Exit Sub

    Dim i As Long
    For i = 0 To UBound(vDocs)
        Dim swDoc As SldWorks.ModelDoc2
        Set swDoc = vDocs(i)

        If Not swDoc Is Nothing Then
            Dim thisPath As String
            thisPath = LCase$(swDoc.GetPathName)

            If thisPath <> "" Then
                If Not baselineDict.Exists(thisPath) Then
                    swApp.CloseDoc swDoc.GetTitle
                End If
            End If
        End If
    Next i

    On Error GoTo 0

End Sub

' --- UTILITY FUNCTIONS -------------------------------------------------------

Private Sub AppendElapsed(ByRef partLog As String, ByVal partStart As Single)
    partLog = partLog & "  Elapsed: " & FormatSeconds(ElapsedSeconds(partStart)) & vbCrLf
End Sub

Private Function NormalizeFolder(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolder = folderPath
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    On Error Resume Next
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    On Error GoTo 0
End Sub

Private Function ElapsedSeconds(ByVal startT As Single) As Double
    Dim t As Double
    t = Timer - startT
    If t < 0 Then t = t + 86400
    ElapsedSeconds = t
End Function

Private Function FormatSeconds(ByVal totalSec As Double) As String
    Dim h As Long, m As Long, s As Long
    h = Int(totalSec / 3600)
    m = Int((totalSec - h * 3600) / 60)
    s = Int(totalSec - h * 3600 - m * 60)
    FormatSeconds = Format$(h, "00") & ":" & Format$(m, "00") & ":" & Format$(s, "00")
End Function