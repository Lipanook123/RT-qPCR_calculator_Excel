' Public quality control thresholds - defined once, used everywhere
Public Const SLOPE_MIN As Double = -3.6
Public Const SLOPE_MAX As Double = -3.1
Public Const R_SQUARED_MIN As Double = 0.98
Public Const INHIBITION_THRESHOLD As Double = 75
Public Const RECOVERY_THRESHOLD As Double = 1

Public Const INITIAL_SAMPLE_VOL = 40
Public Const CONCENTRATED_VOL = 15
Public Const PROCESS_CONTROL_VOL = 0.1
Public Const ELUTION_VOL = 100
Public Const TEMPLATE_VOL = 5




Const BUTTON_CAPTION = "RT-qPCR Calculator"

' Dynamic target detection structure
Type TargetInfo
    Name As String
    IsRecoveryControl As Boolean
    HasStandardCurve As Boolean
    slope As Double
    rSquared As Double
    intercept As Double
    curveValid As Boolean
End Type



'========================================
'====  CREATE BUTTONS ON COMMAND BAR ====
'========================================


Sub Auto_Open()
    '------------VARIABLES----------
    Dim CmdBar As CommandBar
    Dim CmdBarMenu As CommandBarControl
    Dim CmdBarMenuItem As CommandBarControl
    '---------------------------------
    
    Set CmdBar = Application.CommandBars("Worksheet Menu Bar")
    Set CmdBarMenu = CmdBar.Controls("Tools")
    
    ' Remove existing button if it exists
    On Error Resume Next
        Application.DisplayAlerts = False
        CmdBarMenu.Controls(BUTTON_CAPTION).Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Add the qPCR Calculator button
    Set CmdBarMenuItem = CmdBarMenu.Controls.Add(Type:=msoControlButton)
    With CmdBarMenuItem
         .Caption = BUTTON_CAPTION
         .OnAction = "ProcessN1_qPCRData"
    End With
    
End Sub

Sub Auto_Close()
    Dim CmdBar As CommandBar
    Dim CmdBarMenu As CommandBarControl
    
    Set CmdBar = Application.CommandBars("Worksheet Menu Bar")
    Set CmdBarMenu = CmdBar.Controls("Tools")
    
    ' Remove the button when closing
    On Error Resume Next
        Application.DisplayAlerts = False
        CmdBarMenu.Controls(BUTTON_CAPTION).Delete
        Application.DisplayAlerts = True
    On Error GoTo 0
    
End Sub


Function CleanUserName(rawName As String) As String
    ' Remove organization name in parentheses from username
    ' E.g. "David Walker (Cefas)" -> "David Walker"
    
    Dim cleanName As String
    Dim openParenPos As Integer
    
    cleanName = Trim(rawName)
    
    ' Find the position of the opening parenthesis
    openParenPos = InStr(cleanName, "(")
    
    If openParenPos > 0 Then
        ' Remove everything from the opening parenthesis onwards
        cleanName = Trim(Left(cleanName, openParenPos - 1))
    End If
    
    CleanUserName = cleanName
End Function




Sub ProcessN1_qPCRData()
    
    ' Enhanced qPCR Data Analysis with Dynamic Target Detection
    ' Creates new workbook with proper naming convention and file management
    
    Dim ws As Worksheet
    Dim rawWs As Worksheet
    Dim resultsWs As Worksheet
    Dim newWorkbook As Workbook
    Dim lastRow As Long
    Dim dataStartRow As Long
    Dim i As Long
    
    ' Processing parameters (will be set by user form)
    Dim initialSampleVolume As Double
    Dim concentratedVolume As Double
    Dim processControlVolume As Double
    Dim extractionElutionVolume As Double
    Dim qpcrTemplateVolume As Double
    Dim analystName As String
    Dim analystInitials As String
    
    ' Column indices
    Dim colWell As Long, colSampleName As Long, colTargetName As Long
    Dim colTask As Long, colCq As Long, colCqMean As Long, colCqStd As Long
    Dim colQuantity As Long, colQuantityMean As Long, colQuantityStd As Long
    Dim colSlope As Long, colRSquared As Long, colIntercept As Long
    
    ' Dynamic target detection
    Dim detectedTargets() As TargetInfo
    Dim recoveryControlTarget As String
    
    ' Check if there's an active worksheet
    If ActiveSheet Is Nothing Then
        MsgBox "Please open the CSV file first.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    ' Find the data start row
    dataStartRow = 0
    For i = 1 To 10
        If InStr(1, ws.Cells(i, 1).Value, "Well") > 0 Then
            dataStartRow = i
            Exit For
        End If
    Next i
    
    If dataStartRow = 0 Then
        MsgBox "Could not find data headers. Please ensure this is a QuantStudio CSV file.", vbExclamation
        Exit Sub
    End If
    
    ' Show user form to get processing parameters
    If Not ShowParameterForm(initialSampleVolume, concentratedVolume, processControlVolume, _
                           extractionElutionVolume, qpcrTemplateVolume, analystName) Then
        Exit Sub ' User cancelled
    End If
    
    ' Get analyst initials for filename
    analystInitials = GetAnalystInitials(analystName)
    
    ' Find column indices
    Call FindColumnIndices(ws, dataStartRow, colWell, colSampleName, colTargetName, _
                          colTask, colCq, colCqMean, colCqStd, colQuantity, _
                          colQuantityMean, colQuantityStd, colSlope, colRSquared, colIntercept)
    
    ' Check required columns
    If colSampleName = 0 Or colTargetName = 0 Or colQuantityMean = 0 Then
        MsgBox "Required columns not found. Please check your CSV format.", vbExclamation
        Exit Sub
    End If
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, colSampleName).End(xlUp).Row
    
    ' Detect targets automatically and determine recovery control
    Call DetectTargetsAndRecoveryControl(ws, dataStartRow, lastRow, colTargetName, colSlope, colRSquared, colIntercept, _
                                        detectedTargets, recoveryControlTarget)
    
    ' Create new workbook with proper structure
    Set newWorkbook = CreateAnalysisWorkbook(ws, analystInitials)
    Set rawWs = newWorkbook.Worksheets("Raw_CSV")
    Set resultsWs = newWorkbook.Worksheets("Results")
    
    ' Copy raw data to Raw_CSV worksheet
    Call CopyRawData(ws, rawWs)
    
    ' Extract metadata and add analysis info
    Call ExtractMetadata(ws, resultsWs, analystName)
    
    ' Check for positive NTCs
    Dim ntcIssues As String
    ntcIssues = CheckNTCs(ws, dataStartRow, lastRow, colSampleName, colTargetName, colTask, colCqMean)
    
    ' Process samples with dynamic QC
    Call ProcessSamplesWithDynamicQC(ws, resultsWs, dataStartRow, lastRow, colSampleName, colTargetName, _
                                    colTask, colCqMean, colQuantityMean, colQuantityStd, _
                                    detectedTargets, recoveryControlTarget, _
                                    initialSampleVolume, concentratedVolume, processControlVolume, _
                                    extractionElutionVolume, qpcrTemplateVolume, _
                                    INHIBITION_THRESHOLD, RECOVERY_THRESHOLD, ntcIssues)
    
    ' Format results worksheet
    Call FormatResultsWorksheet(resultsWs)
    
    ' Generate filename with detected targets and save as temporary file
    Dim fileName As String
    fileName = GenerateFileNameFromTargets(analystInitials, detectedTargets)
    
    ' Save the workbook as a temporary file
    Call SaveAsTemporary(newWorkbook, fileName)
    
    ' Activate results worksheet
    resultsWs.Activate
    
    ' Show completion message with instructions
    MsgBox "Analysis complete!" & vbCrLf & vbCrLf & _
           "A new workbook has been created with:" & vbCrLf & _
           "- Raw_CSV sheet (original data)" & vbCrLf & _
           "- Results sheet (analysis results)" & vbCrLf & vbCrLf & _
           "Detected targets: " & GetTargetSummary(detectedTargets) & vbCrLf & _
           "Recovery control: " & recoveryControlTarget & vbCrLf & vbCrLf & _
           "The file has been saved as a temporary file." & vbCrLf & _
           "Please use 'Save As' to save it to your desired location." & vbCrLf & vbCrLf & _
           "Suggested filename: " & fileName, vbInformation, "Analysis Complete"
    
End Sub

Sub DetectTargetsAndRecoveryControl(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                                   colTargetName As Long, colSlope As Long, colRSquared As Long, colIntercept As Long, _
                                   ByRef detectedTargets() As TargetInfo, ByRef recoveryControlTarget As String)
    
    ' Detect all unique targets and their properties
    Dim i As Long
    Dim targetName As String
    Dim uniqueTargets As Collection
    Dim target As TargetInfo
    Dim targetCount As Integer
    Dim hasMengo As Boolean
    
    Set uniqueTargets = New Collection
    hasMengo = False
    
    ' First pass: collect unique targets
    For i = dataStartRow + 1 To lastRow
        targetName = UCase(Trim(ws.Cells(i, colTargetName).Value))
        
        If targetName <> "" Then
            ' Try to add to collection (will fail silently if already exists)
            On Error Resume Next
            uniqueTargets.Add targetName, targetName
            On Error GoTo 0
            
            ' Check for Mengo
            If targetName = "MENGO" Then hasMengo = True
        End If
    Next i
    
    ' Convert collection to array and analyse each target
    targetCount = uniqueTargets.Count
    ReDim detectedTargets(1 To targetCount)
    
    For i = 1 To uniqueTargets.Count
        targetName = uniqueTargets(i)
        
        ' Initialize target info
        detectedTargets(i).Name = targetName
        detectedTargets(i).IsRecoveryControl = False
        detectedTargets(i).HasStandardCurve = False
        
        ' Check for standard curve parameters
        Call GetStandardCurveParams(ws, dataStartRow, lastRow, colSlope, colRSquared, colIntercept, _
                                   targetName, detectedTargets(i).slope, detectedTargets(i).rSquared, _
                                   detectedTargets(i).intercept, detectedTargets(i).curveValid, _
                                   SLOPE_MIN, SLOPE_MAX, R_SQUARED_MIN)
        
        If detectedTargets(i).slope <> 0 Then
            detectedTargets(i).HasStandardCurve = True
        End If
    Next i
    
    ' Determine recovery control target
    If hasMengo Then
        recoveryControlTarget = "MENGO"
        ' Mark Mengo as recovery control
        For i = 1 To targetCount
            If detectedTargets(i).Name = "MENGO" Then
                detectedTargets(i).IsRecoveryControl = True
                Exit For
            End If
        Next i
    Else
        ' Ask user about recovery control
        recoveryControlTarget = AskUserForRecoveryControl(detectedTargets)
        
        ' Mark the selected target as recovery control
        For i = 1 To targetCount
            If UCase(detectedTargets(i).Name) = UCase(recoveryControlTarget) Then
                detectedTargets(i).IsRecoveryControl = True
                Exit For
            End If
        Next i
    End If
    
End Sub

Function AskUserForRecoveryControl(detectedTargets() As TargetInfo) As String
    
    ' Ask user if a recovery control was included and which target it is
    Dim response As Integer
    Dim targetList As String
    Dim selectedTarget As String
    Dim i As Integer
    
    ' Build list of detected targets (excluding obvious analytical targets)
    targetList = ""
    For i = 1 To UBound(detectedTargets)
        If detectedTargets(i).Name <> "N1" And detectedTargets(i).Name <> "N2" And _
           Not (InStr(detectedTargets(i).Name, "EC") > 0) Then
            If targetList <> "" Then targetList = targetList & ", "
            targetList = targetList & detectedTargets(i).Name
        End If
    Next i
    
    ' Ask if recovery control was included
    response = MsgBox("Mengo virus recovery control not detected." & vbCrLf & vbCrLf & _
                     "Was a different recovery control included in this assay?" & vbCrLf & vbCrLf & _
                     "Detected targets: " & targetList, vbYesNo + vbQuestion, "Recovery Control")
    
    If response = vbYes Then
        ' Ask which target is the recovery control
        selectedTarget = InputBox("Which target is the recovery control?" & vbCrLf & vbCrLf & _
                                 "Available targets:" & vbCrLf & targetList, _
                                 "Select Recovery Control Target", "")
        
        ' Validate selection
        If selectedTarget <> "" Then
            For i = 1 To UBound(detectedTargets)
                If UCase(detectedTargets(i).Name) = UCase(selectedTarget) Then
                    AskUserForRecoveryControl = UCase(selectedTarget)
                    Exit Function
                End If
            Next i
            
            ' If not found, show error and ask again
            MsgBox "Target '" & selectedTarget & "' not found. Please check spelling.", vbExclamation
            AskUserForRecoveryControl = AskUserForRecoveryControl(detectedTargets)
        Else
            AskUserForRecoveryControl = ""
        End If
    Else
        AskUserForRecoveryControl = ""
    End If
    
End Function

Function GetTargetSummary(detectedTargets() As TargetInfo) As String
    
    ' Create a summary string of detected targets
    Dim summary As String
    Dim i As Integer
    
    summary = ""
    For i = 1 To UBound(detectedTargets)
        If summary <> "" Then summary = summary & ", "
        summary = summary & detectedTargets(i).Name
        If detectedTargets(i).IsRecoveryControl Then summary = summary & " (Recovery)"
    Next i
    
    GetTargetSummary = summary
    
End Function

Function GenerateFileNameFromTargets(analystInitials As String, detectedTargets() As TargetInfo) As String
    
    ' Generate filename from detected targets
    Dim dateStr As String
    Dim targetStr As String
    Dim fileName As String
    Dim i As Integer
    
    dateStr = Format(Date, "yyyymmdd")
    targetStr = ""
    
    ' Build target string from detected targets (excluding EC-RNA controls)
    For i = 1 To UBound(detectedTargets)
        If Not (InStr(detectedTargets(i).Name, "EC") > 0) Then
            If targetStr <> "" Then targetStr = targetStr & "_"
            targetStr = targetStr & detectedTargets(i).Name
        End If
    Next i
    
    If targetStr = "" Then targetStr = "Unknown"
    
    fileName = dateStr & "-" & analystInitials & "-" & targetStr & ".xlsx"
    GenerateFileNameFromTargets = fileName
    
End Function

Sub ProcessSamplesWithDynamicQC(ws As Worksheet, resultsWs As Worksheet, dataStartRow As Long, lastRow As Long, _
                               colSampleName As Long, colTargetName As Long, colTask As Long, colCqMean As Long, _
                               colQuantityMean As Long, colQuantityStd As Long, _
                               detectedTargets() As TargetInfo, recoveryControlTarget As String, _
                               initialSampleVolume As Double, concentratedVolume As Double, processControlVolume As Double, _
                               extractionElutionVolume As Double, qpcrTemplateVolume As Double, _
                               INHIBITION_THRESHOLD, RECOVERY_THRESHOLD, ntcIssues As String)
    
    Dim i As Long
    Dim currentRow As Long
    Dim sampleName As String
    Dim baseSampleName As String
    Dim targetName As String
    Dim task As String
    Dim quantityMean As Double
    Dim quantityStd As Double
    Dim cqMean As Double
    Dim dilutionFactor As Double
    
    ' Conversion factors
    Dim rnaConcentrationFactor As Double
    Dim sampleConcentrationFactor As Double
    Dim totalConcentrationFactor As Double
    
    rnaConcentrationFactor = extractionElutionVolume / qpcrTemplateVolume
    sampleConcentrationFactor = 1000 / concentratedVolume
    totalConcentrationFactor = rnaConcentrationFactor * sampleConcentrationFactor
    
    ' Collections to group samples
    Dim uniqueSamples As Collection
    Set uniqueSamples = New Collection
    
    ' Find primary analytical target (first non-control target)
    Dim primaryTarget As String
    primaryTarget = ""
    For i = 1 To UBound(detectedTargets)
        If Not detectedTargets(i).IsRecoveryControl And Not (InStr(detectedTargets(i).Name, "EC") > 0) Then
            primaryTarget = detectedTargets(i).Name
            Exit For
        End If
    Next i
    
    ' First pass: collect unique samples based on primary target
    For i = dataStartRow + 1 To lastRow
        sampleName = ws.Cells(i, colSampleName).Value
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        task = UCase(ws.Cells(i, colTask).Value)
        
        ' Skip empty rows, standards, and controls; focus on primary target
        If sampleName <> "" And Not IsControl(sampleName, task) And targetName = primaryTarget Then
            baseSampleName = GetBaseSampleName(sampleName)
            
            ' Try to add to unique samples collection
            On Error Resume Next
            uniqueSamples.Add baseSampleName, baseSampleName
            On Error GoTo 0
        End If
    Next i
    
    ' Add standard curve QC section
    currentRow = resultsWs.Cells(resultsWs.Rows.Count, 1).End(xlUp).Row + 2
    resultsWs.Cells(currentRow, 1).Value = "Standard Curve Quality Control:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    ' Display QC for each target with standard curves
    For i = 1 To UBound(detectedTargets)
        If detectedTargets(i).HasStandardCurve Then
            resultsWs.Cells(currentRow, 1).Value = detectedTargets(i).Name & " Standard Curve:"
            resultsWs.Cells(currentRow, 1).Font.Bold = True
            currentRow = currentRow + 1
            
            resultsWs.Cells(currentRow, 1).Value = "Slope:"
            resultsWs.Cells(currentRow, 2).Value = Format(detectedTargets(i).slope, "0.000")
            resultsWs.Cells(currentRow, 3).Value = "Required: " & SLOPE_MIN & " to " & SLOPE_MAX
            resultsWs.Cells(currentRow, 4).Value = IIf(detectedTargets(i).slope >= SLOPE_MIN And detectedTargets(i).slope <= SLOPE_MAX, "PASS", "FAIL")
            currentRow = currentRow + 1
            
            resultsWs.Cells(currentRow, 1).Value = "R²:"
            resultsWs.Cells(currentRow, 2).Value = Format(detectedTargets(i).rSquared, "0.000")
            resultsWs.Cells(currentRow, 3).Value = "Required: >=" & R_SQUARED_MIN
            resultsWs.Cells(currentRow, 4).Value = IIf(detectedTargets(i).rSquared >= R_SQUARED_MIN, "PASS", "FAIL")
            currentRow = currentRow + 1
            
            resultsWs.Cells(currentRow, 1).Value = "Overall Status:"
            resultsWs.Cells(currentRow, 2).Value = IIf(detectedTargets(i).curveValid, "VALID", "INVALID")
            currentRow = currentRow + 2
        End If
    Next i
    
    ' NTC Status
    If ntcIssues <> "" Then
        resultsWs.Cells(currentRow, 1).Value = "NTC Issues:"
        resultsWs.Cells(currentRow, 1).Font.Bold = True
        resultsWs.Cells(currentRow, 1).Font.Color = RGB(255, 0, 0)
        currentRow = currentRow + 1
        resultsWs.Cells(currentRow, 1).Value = ntcIssues
        currentRow = currentRow + 1
    Else
        resultsWs.Cells(currentRow, 1).Value = "NTC Status: PASS (No amplification detected)"
        resultsWs.Cells(currentRow, 1).Font.Color = RGB(0, 128, 0)
        currentRow = currentRow + 1
    End If
    currentRow = currentRow + 1
    
    ' Processing parameters
    resultsWs.Cells(currentRow, 1).Value = "Processing Parameters:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Initial Sample Volume (ml):"
    resultsWs.Cells(currentRow, 2).Value = initialSampleVolume
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Concentrated Volume (ml):"
    resultsWs.Cells(currentRow, 2).Value = concentratedVolume
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "RNA Elution Volume (ul):"
    resultsWs.Cells(currentRow, 2).Value = extractionElutionVolume
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "qPCR Template Volume (ul):"
    resultsWs.Cells(currentRow, 2).Value = qpcrTemplateVolume
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Base Concentration Factor:"
    resultsWs.Cells(currentRow, 2).Value = Format(totalConcentrationFactor, "0.0")
    currentRow = currentRow + 3
    
    ' Create results table header
    resultsWs.Cells(currentRow, 1).Value = "Sample Results (" & primaryTarget & "):"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    ' Table headers - adjust based on whether recovery control is available
    resultsWs.Cells(currentRow, 1).Value = "Sample Name"
    resultsWs.Cells(currentRow, 2).Value = "Dilution Factor"
    resultsWs.Cells(currentRow, 3).Value = "Cq Mean"
    resultsWs.Cells(currentRow, 4).Value = "Quantity Mean"
    resultsWs.Cells(currentRow, 5).Value = "Final Conc. (copies/L)"
    
    Dim colIndex As Integer
    colIndex = 6
    
    ' Add RT-PCR inhibition column if EC-RNA is present
    Dim hasECRNA As Boolean
    hasECRNA = HasECRNATarget(detectedTargets)
    If hasECRNA Then
        resultsWs.Cells(currentRow, colIndex).Value = "RT-PCR Inhibition (%)"
        colIndex = colIndex + 1
    End If
    
    ' Add recovery column if recovery control is present
    If recoveryControlTarget <> "" Then
        resultsWs.Cells(currentRow, colIndex).Value = recoveryControlTarget & " Recovery (%)"
        colIndex = colIndex + 1
    End If
    
    resultsWs.Cells(currentRow, colIndex).Value = "QC Status"
    resultsWs.Cells(currentRow, colIndex + 1).Value = "Calculation Details"
    
    ' Make headers bold
    For i = 1 To colIndex + 1
        resultsWs.Cells(currentRow, i).Font.Bold = True
    Next i
    currentRow = currentRow + 1
    
    ' Process each unique sample
    Dim j As Integer
    For j = 1 To uniqueSamples.Count
        baseSampleName = uniqueSamples(j)
        
        ' Find the best dilution for this sample
        Dim bestSampleName As String
        Dim bestDilution As Double
        Dim bestQuantity As Double
        Dim bestCq As Double
        
        Call FindBestSampleData(ws, dataStartRow, lastRow, colSampleName, colTargetName, colQuantityMean, colCqMean, _
                               baseSampleName, primaryTarget, bestSampleName, bestDilution, bestQuantity, bestCq)
        
        ' Calculate concentrations and QC
        If bestSampleName <> "" Then
            Dim finalConcentration As Double
            finalConcentration = bestQuantity * totalConcentrationFactor * bestDilution
            
            ' Calculate RT-PCR inhibition if EC-RNA is available
            Dim inhibitionPercent As String
            Dim inhibitionValid As Boolean
            inhibitionValid = True ' Default to true if no inhibition test
            
            If hasECRNA Then
                inhibitionPercent = CalculateInhibition(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean, _
                                                       baseSampleName, detectedTargets, inhibitionValid, INHIBITION_THRESHOLD)
            Else
                inhibitionPercent = "Not Available"
            End If
            
            ' Calculate recovery if recovery control is available
            Dim recoveryPercent As String
            Dim recoveryValid As Boolean
            recoveryValid = True ' Default to true if no recovery test
            
            If recoveryControlTarget <> "" Then
                recoveryPercent = CalculateRecovery(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean, _
                                                   baseSampleName, recoveryControlTarget, detectedTargets, _
                                                   concentratedVolume, processControlVolume, recoveryValid, RECOVERY_THRESHOLD)
            Else
                recoveryPercent = "Not Available"
            End If
            
            ' Determine QC status
            Dim qcStatus As String
            qcStatus = DetermineQCStatus(detectedTargets, primaryTarget, recoveryControlTarget, ntcIssues, _
                                       inhibitionValid, recoveryValid, hasECRNA)
            
            ' Calculate details
            Dim calcDetails As String
            calcDetails = Format(bestQuantity, "#,##0") & " x " & Format(totalConcentrationFactor, "0.0") & " x " & bestDilution & " = " & Format(finalConcentration, "#,##0")
            
            ' Write results
            colIndex = 1
            resultsWs.Cells(currentRow, colIndex).Value = baseSampleName: colIndex = colIndex + 1
            resultsWs.Cells(currentRow, colIndex).Value = IIf(bestDilution = 1, "None", "1:" & bestDilution): colIndex = colIndex + 1
            resultsWs.Cells(currentRow, colIndex).Value = IIf(bestCq > 0, Format(bestCq, "0.000"), "N/A"): colIndex = colIndex + 1
            resultsWs.Cells(currentRow, colIndex).Value = Format(bestQuantity, "#,##0"): colIndex = colIndex + 1
            resultsWs.Cells(currentRow, colIndex).Value = Format(finalConcentration, "#,##0"): colIndex = colIndex + 1
            
            If hasECRNA Then
                resultsWs.Cells(currentRow, colIndex).Value = inhibitionPercent: colIndex = colIndex + 1
            End If
            
            If recoveryControlTarget <> "" Then
                resultsWs.Cells(currentRow, colIndex).Value = recoveryPercent: colIndex = colIndex + 1
            End If
            
            resultsWs.Cells(currentRow, colIndex).Value = qcStatus: colIndex = colIndex + 1
            resultsWs.Cells(currentRow, colIndex).Value = calcDetails
            
            currentRow = currentRow + 1
        End If
    Next j
    
    ' Add notes
    Call AddQCNotes(resultsWs, currentRow, hasECRNA, recoveryControlTarget, INHIBITION_THRESHOLD, RECOVERY_THRESHOLD, SLOPE_MIN, SLOPE_MAX, R_SQUARED_MIN)
    
End Sub

Function HasECRNATarget(detectedTargets() As TargetInfo) As Boolean
    
    Dim i As Integer
    For i = 1 To UBound(detectedTargets)
        If InStr(detectedTargets(i).Name, "EC") > 0 Then
            HasECRNATarget = True
            Exit Function
        End If
    Next i
    HasECRNATarget = False
    
End Function

Sub FindBestSampleData(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                      colSampleName As Long, colTargetName As Long, colQuantityMean As Long, colCqMean As Long, _
                      baseSampleName As String, targetName As String, _
                      ByRef bestSampleName As String, ByRef bestDilution As Double, _
                      ByRef bestQuantity As Double, ByRef bestCq As Double)
    
    Dim i As Long
    Dim sampleName As String
    Dim currentTarget As String
    Dim dilutionFactor As Double
    Dim quantityMean As Double
    
    bestDilution = 999999 ' Start with very high number
    bestSampleName = ""
    bestQuantity = 0
    bestCq = 0
    
    For i = dataStartRow + 1 To lastRow
        sampleName = ws.Cells(i, colSampleName).Value
        currentTarget = UCase(ws.Cells(i, colTargetName).Value)
        
        If GetBaseSampleName(sampleName) = baseSampleName And currentTarget = targetName Then
            dilutionFactor = GetDilutionFactor(sampleName)
            quantityMean = 0
            If Not IsEmpty(ws.Cells(i, colQuantityMean).Value) And ws.Cells(i, colQuantityMean).Value <> "" Then
                quantityMean = CDbl(ws.Cells(i, colQuantityMean).Value)
            End If
            
            ' Use this dilution if it's lower and has valid data
            If dilutionFactor < bestDilution And quantityMean > 0 Then
                bestDilution = dilutionFactor
                bestSampleName = sampleName
                bestQuantity = quantityMean
                If Not IsEmpty(ws.Cells(i, colCqMean).Value) Then
                    bestCq = CDbl(ws.Cells(i, colCqMean).Value)
                End If
            End If
        End If
    Next i
    
End Sub

Function CalculateInhibition(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                           colSampleName As Long, colTargetName As Long, colCqMean As Long, _
                           baseSampleName As String, detectedTargets() As TargetInfo, _
                           ByRef inhibitionValid As Boolean, INHIBITION_THRESHOLD) As String
    
    ' Find the primary analytical target slope for inhibition calculation
    Dim primarySlope As Double
    Dim i As Integer
    
    primarySlope = 0
    For i = 1 To UBound(detectedTargets)
        If Not detectedTargets(i).IsRecoveryControl And Not (InStr(detectedTargets(i).Name, "EC") > 0) And detectedTargets(i).HasStandardCurve Then
            primarySlope = detectedTargets(i).slope
            Exit For
        End If
    Next i
    
    If primarySlope = 0 Then
        CalculateInhibition = "No Standard Curve"
        inhibitionValid = False
        Exit Function
    End If
    
    Dim sampleEcCq As Double, controlEcCq As Double, deltaCq As Double, inhibition As Double
    
    sampleEcCq = GetSampleEcRnaCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean, baseSampleName)
    controlEcCq = GetControlEcRnaCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean)
    
    If sampleEcCq > 0 And controlEcCq > 0 Then
        deltaCq = sampleEcCq - controlEcCq
        inhibition = (1 - (10 ^ (deltaCq / primarySlope))) * 100
        inhibitionValid = (inhibition < INHIBITION_THRESHOLD)
        CalculateInhibition = Format(inhibition, "0.1") & "%"
    Else
        CalculateInhibition = "Not Calculated"
        inhibitionValid = False
    End If
    
End Function

Function CalculateRecovery(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                         colSampleName As Long, colTargetName As Long, colCqMean As Long, _
                         baseSampleName As String, recoveryControlTarget As String, detectedTargets() As TargetInfo, _
                         concentratedVolume As Double, processControlVolume As Double, _
                         ByRef recoveryValid As Boolean, RECOVERY_THRESHOLD) As String
    
    ' Find the recovery control target slope
    Dim recoverySlope As Double
    Dim i As Integer
    
    recoverySlope = 0
    For i = 1 To UBound(detectedTargets)
        If detectedTargets(i).Name = recoveryControlTarget And detectedTargets(i).HasStandardCurve Then
            recoverySlope = detectedTargets(i).slope
            Exit For
        End If
    Next i
    
    If recoverySlope = 0 Then
        CalculateRecovery = "No Standard Curve"
        recoveryValid = False
        Exit Function
    End If
    
    Dim sampleRecoveryCq As Double, processControlRecoveryCq As Double, deltaCq As Double
    Dim samplePathDilution As Double, expectedDilutionDifference As Double
    Dim rawRecovery As Double, recovery As Double
    
    sampleRecoveryCq = GetSampleRecoveryControlCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean, baseSampleName, recoveryControlTarget)
    processControlRecoveryCq = GetProcessControlRecoveryControlCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean, recoveryControlTarget)
    
    If sampleRecoveryCq > 0 And processControlRecoveryCq > 0 Then
        deltaCq = sampleRecoveryCq - processControlRecoveryCq
        samplePathDilution = concentratedVolume / processControlVolume
        expectedDilutionDifference = samplePathDilution / 1 ' Control is direct extraction
        
        rawRecovery = (10 ^ (deltaCq / recoverySlope)) * 100
        recovery = rawRecovery * expectedDilutionDifference
        recoveryValid = (recovery > RECOVERY_THRESHOLD)
        CalculateRecovery = Format(recovery, "0.1") & "%"
    Else
        CalculateRecovery = "Not Calculated"
        recoveryValid = False
    End If
    
End Function

Function DetermineQCStatus(detectedTargets() As TargetInfo, primaryTarget As String, recoveryControlTarget As String, _
                         ntcIssues As String, inhibitionValid As Boolean, recoveryValid As Boolean, hasECRNA As Boolean) As String
    
    ' Find primary target curve validity
    Dim primaryCurveValid As Boolean
    Dim recoveryCurveValid As Boolean
    Dim i As Integer
    
    primaryCurveValid = False
    recoveryCurveValid = True ' Default to true if no recovery control
    
    For i = 1 To UBound(detectedTargets)
        If detectedTargets(i).Name = primaryTarget And detectedTargets(i).HasStandardCurve Then
            primaryCurveValid = detectedTargets(i).curveValid
        End If
        If detectedTargets(i).Name = recoveryControlTarget And detectedTargets(i).HasStandardCurve Then
            recoveryCurveValid = detectedTargets(i).curveValid
        End If
    Next i
    
    ' Determine overall status
    If primaryCurveValid And recoveryCurveValid And ntcIssues = "" Then
        If hasECRNA And recoveryControlTarget <> "" Then
            ' Full QC available
            If inhibitionValid And recoveryValid Then
                DetermineQCStatus = "Valid"
            Else
                DetermineQCStatus = "QC Fail"
            End If
        ElseIf hasECRNA Then
            ' Only inhibition QC
            If inhibitionValid Then
                DetermineQCStatus = "Valid (No Recovery QC)"
            Else
                DetermineQCStatus = "Inhibition Fail"
            End If
        ElseIf recoveryControlTarget <> "" Then
            ' Only recovery QC
            If recoveryValid Then
                DetermineQCStatus = "Valid (No Inhibition QC)"
            Else
                DetermineQCStatus = "Recovery Fail"
            End If
        Else
            ' No QC available
            DetermineQCStatus = "Valid (No QC Available)"
        End If
    Else
        DetermineQCStatus = "Invalid"
    End If
    
End Function

Function GetSampleRecoveryControlCq(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                                  colSampleName As Long, colTargetName As Long, colCqMean As Long, _
                                  baseSampleName As String, recoveryControlTarget As String) As Double
    
    Dim i As Long
    Dim sampleName As String
    Dim targetName As String
    Dim cqValue As Double
    
    For i = dataStartRow + 1 To lastRow
        sampleName = ws.Cells(i, colSampleName).Value
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        
        If GetBaseSampleName(sampleName) = baseSampleName And targetName = recoveryControlTarget Then
            If Not IsEmpty(ws.Cells(i, colCqMean).Value) And ws.Cells(i, colCqMean).Value <> "" Then
                cqValue = CDbl(ws.Cells(i, colCqMean).Value)
                GetSampleRecoveryControlCq = cqValue
                Exit Function
            End If
        End If
    Next i
    
    GetSampleRecoveryControlCq = 0 ' Not found
    
End Function

Function GetProcessControlRecoveryControlCq(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                                          colSampleName As Long, colTargetName As Long, colCqMean As Long, _
                                          recoveryControlTarget As String) As Double
    
    Dim i As Long
    Dim sampleName As String
    Dim targetName As String
    Dim cqValue As Double
    
    For i = dataStartRow + 1 To lastRow
        sampleName = UCase(ws.Cells(i, colSampleName).Value)
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        
        If InStr(sampleName, "PROCESS") > 0 And targetName = recoveryControlTarget Then
            If Not IsEmpty(ws.Cells(i, colCqMean).Value) And ws.Cells(i, colCqMean).Value <> "" Then
                cqValue = CDbl(ws.Cells(i, colCqMean).Value)
                GetProcessControlRecoveryControlCq = cqValue
                Exit Function
            End If
        End If
    Next i
    
    GetProcessControlRecoveryControlCq = 0 ' Not found
    
End Function

Sub AddQCNotes(resultsWs As Worksheet, ByRef currentRow As Long, hasECRNA As Boolean, recoveryControlTarget As String, _
               INHIBITION_THRESHOLD, RECOVERY_THRESHOLD, SLOPE_MIN, SLOPE_MAX, R_SQUARED_MIN)
    
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "QC Criteria:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    If hasECRNA Then
        resultsWs.Cells(currentRow, 1).Value = "- RT-qPCR Inhibition: <" & INHIBITION_THRESHOLD & "% acceptable"
        currentRow = currentRow + 1
    End If
    
    If recoveryControlTarget <> "" Then
        resultsWs.Cells(currentRow, 1).Value = "- " & recoveryControlTarget & " Recovery: >" & RECOVERY_THRESHOLD & "% acceptable"
        currentRow = currentRow + 1
    End If
    
    resultsWs.Cells(currentRow, 1).Value = "- Standard Curves: Slope " & SLOPE_MIN & " to " & SLOPE_MAX & ", R² >=" & R_SQUARED_MIN
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "- NTC: No amplification detected"
    currentRow = currentRow + 1
    
    If Not hasECRNA Or recoveryControlTarget = "" Then
        resultsWs.Cells(currentRow, 1).Value = "* Limited QC available - some quality control measures not included in this assay"
        resultsWs.Cells(currentRow, 1).Font.Italic = True
        currentRow = currentRow + 1
    End If
    
End Sub

' [Keep all your existing helper functions unchanged]

Function CreateAnalysisWorkbook(sourceWs As Worksheet, analystInitials As String) As Workbook
    ' Create new workbook with proper structure
    Dim newWb As Workbook
    Dim rawWs As Worksheet
    Dim resultsWs As Worksheet
    
    Set newWb = Workbooks.Add
    
    ' Delete extra sheets and rename/create required sheets
    Application.DisplayAlerts = False
    Do While newWb.Worksheets.Count > 1
        newWb.Worksheets(newWb.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Rename first sheet to Raw_CSV
    newWb.Worksheets(1).Name = "Raw_CSV"
    Set rawWs = newWb.Worksheets("Raw_CSV")
    
    ' Add Results worksheet
    Set resultsWs = newWb.Worksheets.Add(After:=rawWs)
    resultsWs.Name = "Results"
    
    Set CreateAnalysisWorkbook = newWb
End Function

Sub CopyRawData(sourceWs As Worksheet, targetWs As Worksheet)
    ' Copy all raw data from source worksheet to target worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    
    ' Find the extent of data more robustly
    lastRow = 0
    lastCol = 0
    
    ' Find last row with data in any column
    For i = 1 To 1000 ' Check first 1000 rows
        For j = 1 To 50 ' Check first 50 columns
            If sourceWs.Cells(i, j).Value <> "" Then
                If i > lastRow Then lastRow = i
                If j > lastCol Then lastCol = j
            End If
        Next j
    Next i
    
    ' Ensure we have at least some data
    If lastRow = 0 Then lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    If lastCol = 0 Then lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column
    
    ' Copy data using values to avoid formatting issues
    For i = 1 To lastRow
        For j = 1 To lastCol
            targetWs.Cells(i, j).Value = sourceWs.Cells(i, j).Value
        Next j
    Next i
    
    ' Add header information
    targetWs.Cells(1, lastCol + 2).Value = "Original File:"
    targetWs.Cells(2, lastCol + 2).Value = sourceWs.Parent.Name
    targetWs.Cells(3, lastCol + 2).Value = "Copied on:"
    targetWs.Cells(4, lastCol + 2).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Auto-fit columns
    targetWs.Columns.AutoFit
End Sub

Function GetAnalystInitials(analystName As String) As String
    ' Extract initials from analyst name
    Dim initials As String
    Dim words() As String
    Dim i As Integer
    
    ' Clean the name
    analystName = Trim(analystName)
    
    ' Split by spaces
    words = Split(analystName, " ")
    
    initials = ""
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            initials = initials & UCase(Left(words(i), 1))
        End If
    Next i
    
    ' If no spaces found, take first two characters
    If initials = "" And Len(analystName) >= 2 Then
        initials = UCase(Left(analystName, 2))
    ElseIf initials = "" Then
        initials = "XX"
    End If
    
    ' Limit to 3 characters maximum
    If Len(initials) > 3 Then
        initials = Left(initials, 3)
    End If
    
    GetAnalystInitials = initials
End Function

Sub SaveAsTemporary(wb As Workbook, suggestedName As String)
    ' Save workbook as temporary file with suggested name
    Dim tempPath As String
    Dim tempFileName As String
    
    ' Get temporary directory
    tempPath = Environ("TEMP")
    If Right(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
    
    ' Create temporary filename
    tempFileName = tempPath & "TEMP_" & suggestedName
    
    ' Save the workbook
    On Error GoTo SaveError
    wb.SaveAs fileName:=tempFileName, FileFormat:=xlOpenXMLWorkbook
    
    ' Set the suggested name for future saves
    wb.Application.Caption = "Microsoft Excel - " & suggestedName & " (Use Save As to save to desired location)"
    
    Exit Sub
    
SaveError:
    MsgBox "Could not save temporary file. The workbook remains unsaved." & vbCrLf & _
           "Please use Save As to save manually.", vbExclamation
End Sub

Function ShowParameterForm(ByRef initialSampleVol As Double, ByRef concentratedVol As Double, _
                          ByRef processControlVol As Double, ByRef extractionElutionVol As Double, _
                          ByRef qpcrTemplateVol As Double, ByRef analystName As String) As Boolean
    
    ' Create a simple input form using InputBox for parameters
    Dim response As String
    Dim proceed As Boolean
    Dim rawUserName As String
    
    rawUserName = Application.UserName
    
    proceed = True
    
    ' Get analyst name first
    analystName = InputBox("Enter analyst name (for file naming):", "Analyst Information", CleanUserName(rawUserName))
    If analystName = "" Then
        proceed = False
        GoTo ExitFunction
    End If
    
    ' Show initial message about parameters
    If MsgBox("Do you want to use default processing parameters?" & vbCrLf & vbCrLf & _
              "Defaults:" & vbCrLf & _
              "Initial Sample Volume: " & INITIAL_SAMPLE_VOL & " ml" & vbCrLf & _
              "Concentrated Volume: " & CONCENTRATED_VOL & " ml" & vbCrLf & _
              "Process Control Volume: " & PROCESS_CONTROL_VOL & " ml" & vbCrLf & _
              "RNA Elution Volume: " & ELUTION_VOL & " µl" & vbCrLf & _
              "qPCR Template Volume: " & TEMPLATE_VOL & " µl", vbYesNo + vbQuestion, "Processing Parameters") = vbYes Then
        
        ' Use defaults
        initialSampleVol = INITIAL_SAMPLE_VOL
        concentratedVol = CONCENTRATED_VOL
        processControlVol = PROCESS_CONTROL_VOL
        extractionElutionVol = ELUTION_VOL
        qpcrTemplateVol = TEMPLATE_VOL
        
        
        
        
        
    Else
        ' Get custom parameters
        response = InputBox("Enter Initial Sample Volume (ml):", "Processing Parameters", INITIAL_SAMPLE_VOL)
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        initialSampleVol = CDbl(response)
        
        response = InputBox("Enter Concentrated Volume (ml):", "Processing Parameters", CONCENTRATED_VOL)
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        concentratedVol = CDbl(response)
        
        response = InputBox("Enter Process Control Volume (ml):", "Processing Parameters", PROCESS_CONTROL_VOL)
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        processControlVol = CDbl(response)
        
        response = InputBox("Enter RNA Elution Volume (ul):", "Processing Parameters", ELUTION_VOL)
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        extractionElutionVol = CDbl(response)
        
        response = InputBox("Enter qPCR Template Volume (ul):", "Processing Parameters", TEMPLATE_VOL)
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        qpcrTemplateVol = CDbl(response)
    End If

ExitFunction:
    ShowParameterForm = proceed
    
End Function

Sub FindColumnIndices(ws As Worksheet, headerRow As Long, _
                     ByRef colWell As Long, ByRef colSampleName As Long, ByRef colTargetName As Long, _
                     ByRef colTask As Long, ByRef colCq As Long, ByRef colCqMean As Long, ByRef colCqStd As Long, _
                     ByRef colQuantity As Long, ByRef colQuantityMean As Long, ByRef colQuantityStd As Long, _
                     ByRef colSlope As Long, ByRef colRSquared As Long, ByRef colIntercept As Long)
    
    Dim col As Long
    Dim headerText As String
    
    ' Reset all column indices
    colWell = 0: colSampleName = 0: colTargetName = 0: colTask = 0
    colCq = 0: colCqMean = 0: colCqStd = 0: colQuantity = 0
    colQuantityMean = 0: colQuantityStd = 0: colSlope = 0: colRSquared = 0: colIntercept = 0
    
    ' Search for columns
    For col = 1 To 30 ' Limit search to first 30 columns
        headerText = UCase(Replace(ws.Cells(headerRow, col).Value, """", ""))
        
        Select Case headerText
            Case "WELL"
                colWell = col
            Case "SAMPLE NAME"
                colSampleName = col
            Case "TARGET NAME"
                colTargetName = col
            Case "TASK"
                colTask = col
            Case "CQ"
                colCq = col
            Case "CQ MEAN"
                colCqMean = col
            Case "CQ STANDARD DEVIATION"
                colCqStd = col
            Case "QUANTITY"
                colQuantity = col
            Case "QUANTITY MEAN"
                colQuantityMean = col
            Case "QUANTITY STANDARD DEVIATION"
                colQuantityStd = col
            Case "SLOPE"
                colSlope = col
            Case "R-SQUARED"
                colRSquared = col
            Case "INTERCEPT"
                colIntercept = col
        End Select
    Next col
    
    ' Use fallbacks for missing columns
    If colCqMean = 0 And colCq > 0 Then colCqMean = colCq
    If colQuantityMean = 0 And colQuantity > 0 Then colQuantityMean = colQuantity
    
End Sub

Sub ExtractMetadata(ws As Worksheet, resultsWs As Worksheet, analystName As String)
    
    Dim i As Long
    Dim cellValue As String
    Dim currentRow As Long
    
    currentRow = 1
    
    ' Add title
    resultsWs.Cells(currentRow, 1).Value = "qPCR Analysis Results"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    resultsWs.Cells(currentRow, 1).Font.Size = 14
    currentRow = currentRow + 2
    
    ' Extract metadata from first few rows
    For i = 1 To 10
        cellValue = ws.Cells(i, 1).Value
        If Left(cellValue, 1) = "#" Then
            resultsWs.Cells(currentRow, 1).Value = Replace(cellValue, "#", "")
            currentRow = currentRow + 1
        End If
    Next i
    
    ' Add analysis info
    resultsWs.Cells(currentRow, 1).Value = "Analysed By: " & analystName
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "Analysis Date: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    currentRow = currentRow + 2
    
End Sub

Sub GetStandardCurveParams(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                          colSlope As Long, colRSquared As Long, colIntercept As Long, _
                          targetName As String, ByRef slope As Double, ByRef rSquared As Double, _
                          ByRef intercept As Double, ByRef curveValid As Boolean, _
                          SLOPE_MIN, SLOPE_MAX, R_SQUARED_MIN)
    
    Dim i As Long
    Dim foundParams As Boolean
    
    foundParams = False
    slope = 0: rSquared = 0: intercept = 0
    
    ' Find standard curve parameters for specific target
    For i = dataStartRow + 1 To lastRow
        If UCase(ws.Cells(i, 3).Value) = UCase(targetName) Then ' Assuming target name is column 3
            If colSlope > 0 And Not IsEmpty(ws.Cells(i, colSlope).Value) And ws.Cells(i, colSlope).Value <> "" Then
                slope = CDbl(ws.Cells(i, colSlope).Value)
                foundParams = True
            End If
            If colRSquared > 0 And Not IsEmpty(ws.Cells(i, colRSquared).Value) And ws.Cells(i, colRSquared).Value <> "" Then
                rSquared = CDbl(ws.Cells(i, colRSquared).Value)
            End If
            If colIntercept > 0 And Not IsEmpty(ws.Cells(i, colIntercept).Value) And ws.Cells(i, colIntercept).Value <> "" Then
                intercept = CDbl(ws.Cells(i, colIntercept).Value)
            End If
            If foundParams Then Exit For
        End If
    Next i
    
    ' Validate standard curve
    curveValid = (slope >= SLOPE_MIN And slope <= SLOPE_MAX And rSquared >= R_SQUARED_MIN)
    
End Sub

Function CheckNTCs(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                   colSampleName As Long, colTargetName As Long, colTask As Long, colCqMean As Long) As String
    
    Dim i As Long
    Dim sampleName As String
    Dim targetName As String
    Dim task As String
    Dim cqMean As String
    Dim issues As String
    
    issues = ""
    
    For i = dataStartRow + 1 To lastRow
        sampleName = UCase(ws.Cells(i, colSampleName).Value)
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        task = UCase(ws.Cells(i, colTask).Value)
        cqMean = ws.Cells(i, colCqMean).Value
        
        ' Check if it's an NTC with amplification
        If (InStr(sampleName, "NTC") > 0 Or InStr(task, "NTC") > 0) And cqMean <> "" Then
            issues = issues & "POSITIVE NTC detected: " & ws.Cells(i, colSampleName).Value & " (" & targetName & ")" & vbCrLf
        End If
    Next i
    
    CheckNTCs = issues
    
End Function

Function GetDilutionFactor(sampleName As String) As Double
    
    ' Extract dilution factor from sample name
    ' Looks for patterns like "1-10", "1:10", "_10x", etc.
    
    Dim dilutionFactor As Double
    dilutionFactor = 1 ' Default no dilution
    
    ' Check for common dilution patterns
    If InStr(sampleName, "1-10") > 0 Or InStr(sampleName, "1:10") > 0 Then
        dilutionFactor = 10
    ElseIf InStr(sampleName, "1-100") > 0 Or InStr(sampleName, "1:100") > 0 Then
        dilutionFactor = 100
    ElseIf InStr(sampleName, "_10x") > 0 Or InStr(sampleName, "_10X") > 0 Then
        dilutionFactor = 10
    ElseIf InStr(sampleName, "_100x") > 0 Or InStr(sampleName, "_100X") > 0 Then
        dilutionFactor = 100
    End If
    
    GetDilutionFactor = dilutionFactor
    
End Function

Function GetBaseSampleName(sampleName As String) As String
    
    ' Remove dilution indicators to get base sample name for grouping
    Dim baseName As String
    
    baseName = sampleName
    
    ' Remove common dilution suffixes
    baseName = Replace(baseName, "_1-10", "")
    baseName = Replace(baseName, "_1:10", "")
    baseName = Replace(baseName, "_10x", "")
    baseName = Replace(baseName, "_10X", "")
    baseName = Replace(baseName, "_1-100", "")
    baseName = Replace(baseName, "_1:100", "")
    baseName = Replace(baseName, "_100x", "")
    baseName = Replace(baseName, "_100X", "")
    
    GetBaseSampleName = baseName
    
End Function

Function IsControl(sampleName As String, task As String) As Boolean
    
    Dim sampleUpper As String
    Dim taskUpper As String
    
    sampleUpper = UCase(sampleName)
    taskUpper = UCase(task)
    
    ' Check if it's a control or standard
    IsControl = (InStr(sampleUpper, "STD") > 0 Or _
                 InStr(sampleUpper, "STANDARD") > 0 Or _
                 InStr(sampleUpper, "NTC") > 0 Or _
                 InStr(sampleUpper, "H2O") > 0 Or _
                 InStr(sampleUpper, "WATER") > 0 Or _
                 InStr(sampleUpper, "CONTROL") > 0 Or _
                 InStr(sampleUpper, "PROCESS") > 0 Or _
                 InStr(sampleUpper, "NEG") > 0 Or _
                 InStr(taskUpper, "STANDARD") > 0 Or _
                 InStr(taskUpper, "NTC") > 0)
    
End Function

Function GetSampleEcRnaCq(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                         colSampleName As Long, colTargetName As Long, colCqMean As Long, _
                         baseSampleName As String) As Double
    
    ' Find EC-RNA Cq value for the specified sample
    Dim i As Long
    Dim sampleName As String
    Dim targetName As String
    Dim cqValue As Double
    
    For i = dataStartRow + 1 To lastRow
        sampleName = ws.Cells(i, colSampleName).Value
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        
        If GetBaseSampleName(sampleName) = baseSampleName And InStr(targetName, "EC") > 0 Then
            If Not IsEmpty(ws.Cells(i, colCqMean).Value) And ws.Cells(i, colCqMean).Value <> "" Then
                cqValue = CDbl(ws.Cells(i, colCqMean).Value)
                GetSampleEcRnaCq = cqValue
                Exit Function
            End If
        End If
    Next i
    
    GetSampleEcRnaCq = 0 ' Not found
    
End Function

Function GetControlEcRnaCq(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                          colSampleName As Long, colTargetName As Long, colCqMean As Long) As Double
    
    ' Find EC-RNA Cq value for water/negative control
    Dim i As Long
    Dim sampleName As String
    Dim targetName As String
    Dim cqValue As Double
    
    For i = dataStartRow + 1 To lastRow
        sampleName = UCase(ws.Cells(i, colSampleName).Value)
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        
        If (InStr(sampleName, "WATER") > 0 Or InStr(sampleName, "CONTROL") > 0) And InStr(targetName, "EC") > 0 Then
            If Not IsEmpty(ws.Cells(i, colCqMean).Value) And ws.Cells(i, colCqMean).Value <> "" Then
                cqValue = CDbl(ws.Cells(i, colCqMean).Value)
                GetControlEcRnaCq = cqValue
                Exit Function
            End If
        End If
    Next i
    
    GetControlEcRnaCq = 0 ' Not found
    
End Function

Function GetSampleMengoCq(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                         colSampleName As Long, colTargetName As Long, colCqMean As Long, _
                         baseSampleName As String) As Double
    
    ' Find Mengo virus Cq value for the specified sample
    Dim i As Long
    Dim sampleName As String
    Dim targetName As String
    Dim cqValue As Double
    
    For i = dataStartRow + 1 To lastRow
        sampleName = ws.Cells(i, colSampleName).Value
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        
        If GetBaseSampleName(sampleName) = baseSampleName And targetName = "MENGO" Then
            If Not IsEmpty(ws.Cells(i, colCqMean).Value) And ws.Cells(i, colCqMean).Value <> "" Then
                cqValue = CDbl(ws.Cells(i, colCqMean).Value)
                GetSampleMengoCq = cqValue
                Exit Function
            End If
        End If
    Next i
    
    GetSampleMengoCq = 0 ' Not found
    
End Function

Function GetProcessControlMengoCq(ws As Worksheet, dataStartRow As Long, lastRow As Long, _
                                 colSampleName As Long, colTargetName As Long, colCqMean As Long) As Double
    
    ' Find Mengo virus Cq value for process control
    Dim i As Long
    Dim sampleName As String
    Dim targetName As String
    Dim cqValue As Double
    
    For i = dataStartRow + 1 To lastRow
        sampleName = UCase(ws.Cells(i, colSampleName).Value)
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        
        If InStr(sampleName, "PROCESS") > 0 And targetName = "MENGO" Then
            If Not IsEmpty(ws.Cells(i, colCqMean).Value) And ws.Cells(i, colCqMean).Value <> "" Then
                cqValue = CDbl(ws.Cells(i, colCqMean).Value)
                GetProcessControlMengoCq = cqValue
                Exit Function
            End If
        End If
    Next i
    
    GetProcessControlMengoCq = 0 ' Not found
    
End Function

Sub FormatResultsWorksheet(resultsWs As Worksheet)
    
    ' Auto-fit columns
    resultsWs.Columns.AutoFit
    
    ' Find the results table and apply formatting
    Dim tableStartRow As Long
    Dim tableEndRow As Long
    Dim i As Long
    
    ' Find "Sample Results:" header
    For i = 1 To 100
        If InStr(resultsWs.Cells(i, 1).Value, "Sample Results") > 0 Then
            tableStartRow = i + 1 ' Header row
            Exit For
        End If
    Next i
    
    If tableStartRow > 0 Then
        ' Find end of table (look for QC Criteria section or end of data)
        For i = tableStartRow To resultsWs.Cells(resultsWs.Rows.Count, 1).End(xlUp).Row
            If InStr(resultsWs.Cells(i, 1).Value, "QC Criteria") > 0 Then
                tableEndRow = i - 2 ' Stop before QC Criteria
                Exit For
            End If
        Next i
        
        If tableEndRow = 0 Then
            tableEndRow = resultsWs.Cells(resultsWs.Rows.Count, 1).End(xlUp).Row
        End If
        
        ' Apply table formatting
        With resultsWs.Range(resultsWs.Cells(tableStartRow, 1), resultsWs.Cells(tableEndRow, 10))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' Header row formatting
        With resultsWs.Range(resultsWs.Cells(tableStartRow, 1), resultsWs.Cells(tableStartRow, 10))
            .Interior.Color = RGB(200, 200, 200)
            .Font.Bold = True
        End With
        
        ' Color code QC status column (find the column with "QC Status")
        Dim qcColumn As Long
        qcColumn = 0
        For i = 1 To 10
            If InStr(resultsWs.Cells(tableStartRow, i).Value, "QC Status") > 0 Then
                qcColumn = i
                Exit For
            End If
        Next i
        
        If qcColumn > 0 Then
            For i = tableStartRow + 1 To tableEndRow
                If InStr(resultsWs.Cells(i, qcColumn).Value, "Valid") > 0 Then
                    resultsWs.Cells(i, qcColumn).Interior.Color = RGB(200, 255, 200) ' Light green
                ElseIf InStr(resultsWs.Cells(i, qcColumn).Value, "Fail") > 0 Or resultsWs.Cells(i, qcColumn).Value = "Invalid" Then
                    resultsWs.Cells(i, qcColumn).Interior.Color = RGB(255, 200, 200) ' Light red
                ElseIf InStr(resultsWs.Cells(i, qcColumn).Value, "No") > 0 Then
                    resultsWs.Cells(i, qcColumn).Interior.Color = RGB(255, 255, 200) ' Light yellow
                End If
            Next i
        End If
    End If
    
    ' Set column widths for better readability
    resultsWs.Columns("A:A").ColumnWidth = 18 ' Sample Name
    resultsWs.Columns("B:B").ColumnWidth = 12 ' Dilution Factor
    resultsWs.Columns("C:C").ColumnWidth = 10 ' Cq Mean
    resultsWs.Columns("D:D").ColumnWidth = 12 ' Quantity Mean
    resultsWs.Columns("E:E").ColumnWidth = 15 ' Final Concentration
    resultsWs.Columns("F:F").ColumnWidth = 15 ' Inhibition/Recovery columns
    resultsWs.Columns("G:G").ColumnWidth = 15
    resultsWs.Columns("H:H").ColumnWidth = 15
    resultsWs.Columns("I:I").ColumnWidth = 12 ' QC Status
    resultsWs.Columns("J:J").ColumnWidth = 30 ' Calculation Details
    
End Sub
