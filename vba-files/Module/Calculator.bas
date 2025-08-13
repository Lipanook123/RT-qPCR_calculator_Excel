Attribute VB_Name = "Calculator"
' Public quality control thresholds - defined once, used everywhere
Public Const SLOPE_MIN As Double = -3.6
Public Const SLOPE_MAX As Double = -3.1
Public Const R_SQUARED_MIN As Double = 0.98
Public Const INHIBITION_THRESHOLD As Double = 75
Public Const RECOVERY_THRESHOLD As Double = 1




Sub ProcessN1_qPCRData()
    
    ' Enhanced N1 qPCR Data Analysis for QuantStudio CSV Files
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
    
    ' Create new workbook with proper structure
    Set newWorkbook = CreateAnalysisWorkbook(ws, analystInitials)
    Set rawWs = newWorkbook.Worksheets("Raw_CSV")
    Set resultsWs = newWorkbook.Worksheets("Results")
    
    ' Copy raw data to Raw_CSV worksheet
    Call CopyRawData(ws, rawWs)
    
    ' Extract metadata and add analysis info
    Call ExtractMetadata(ws, resultsWs, analystName)
    
    ' Determine which targets are present
    Dim targetsPresent As String
    targetsPresent = DetermineTargets(ws, dataStartRow, lastRow, colTargetName)
    
    ' Get standard curve parameters for N1 and Mengo (if present)
    Dim n1Slope As Double, n1RSquared As Double, n1Intercept As Double, n1CurveValid As Boolean
    Dim mengoSlope As Double, mengoRSquared As Double, mengoIntercept As Double, mengoCurveValid As Boolean
    
    Call GetStandardCurveParams(ws, dataStartRow, lastRow, colSlope, colRSquared, colIntercept, _
                               "N1", n1Slope, n1RSquared, n1Intercept, n1CurveValid, SLOPE_MIN, SLOPE_MAX, R_SQUARED_MIN)
    
    If InStr(targetsPresent, "Mengo") > 0 Then
        Call GetStandardCurveParams(ws, dataStartRow, lastRow, colSlope, colRSquared, colIntercept, _
                                   "Mengo", mengoSlope, mengoRSquared, mengoIntercept, mengoCurveValid, SLOPE_MIN, SLOPE_MAX, R_SQUARED_MIN)
    Else
        mengoCurveValid = False
    End If
    
    ' Check for positive NTCs
    Dim ntcIssues As String
    ntcIssues = CheckNTCs(ws, dataStartRow, lastRow, colSampleName, colTargetName, colTask, colCqMean)
    
    ' Process samples with full QC
    Call ProcessSamplesWithQC(ws, resultsWs, dataStartRow, lastRow, colSampleName, colTargetName, _
                             colTask, colCqMean, colQuantityMean, colQuantityStd, _
                             n1Slope, n1RSquared, n1Intercept, n1CurveValid, _
                             mengoSlope, mengoRSquared, mengoIntercept, mengoCurveValid, _
                             initialSampleVolume, concentratedVolume, processControlVolume, _
                             extractionElutionVolume, qpcrTemplateVolume, _
                             INHIBITION_THRESHOLD, RECOVERY_THRESHOLD, ntcIssues)
    
    ' Format results worksheet
    Call FormatResultsWorksheet(resultsWs)
    
    ' Generate filename and save as temporary file
    Dim fileName As String
    fileName = GenerateFileName(analystInitials, targetsPresent)
    
    ' Save the workbook as a temporary file
    Call SaveAsTemporary(newWorkbook, fileName)
    
    ' Activate results worksheet
    resultsWs.Activate
    
    ' Show completion message with instructions
    MsgBox "Analysis complete!" & vbCrLf & vbCrLf & _
           "A new workbook has been created with:" & vbCrLf & _
           "� Raw_CSV sheet (original data)" & vbCrLf & _
           "� Results sheet (analysis results)" & vbCrLf & vbCrLf & _
           "The file has been saved as a temporary file." & vbCrLf & _
           "Please use 'Save As' to save it to your desired location." & vbCrLf & vbCrLf & _
           "Suggested filename: " & fileName, vbInformation, "Analysis Complete"
    
End Sub

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
    
    ' Find the extent of data
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column
    
    ' Copy data
    sourceWs.Range(sourceWs.Cells(1, 1), sourceWs.Cells(lastRow, lastCol)).Copy
    targetWs.Range("A1").PasteSpecial xlPasteAll
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    ' Add header information
    targetWs.Cells(1, lastCol + 2).Value = "Original File:"
    targetWs.Cells(2, lastCol + 2).Value = sourceWs.Parent.Name
    targetWs.Cells(3, lastCol + 2).Value = "Copied on:"
    targetWs.Cells(4, lastCol + 2).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    ' Auto-fit columns
    targetWs.Columns.AutoFit
End Sub

Function DetermineTargets(ws As Worksheet, dataStartRow As Long, lastRow As Long, colTargetName As Long) As String
    ' Determine which targets are present in the data
    Dim i As Long
    Dim targetName As String
    Dim targets As String
    Dim hasN1 As Boolean
    Dim hasMengo As Boolean
    Dim hasECRNA As Boolean
    
    hasN1 = False
    hasMengo = False
    hasECRNA = False
    
    For i = dataStartRow + 1 To lastRow
        targetName = UCase(Trim(ws.Cells(i, colTargetName).Value))
        
        If targetName = "N1" Then hasN1 = True
        If targetName = "MENGO" Then hasMengo = True
        If InStr(targetName, "EC") > 0 Then hasECRNA = True
    Next i
    
    ' Build target string
    targets = ""
    If hasN1 Then targets = "N1"
    If hasMengo Then
        If targets <> "" Then targets = targets & "_"
        targets = targets & "Mengo"
    End If
    If hasECRNA Then
        If targets <> "" Then targets = targets & "_"
        targets = targets & "EC"
    End If
    
    If targets = "" Then targets = "Unknown"
    
    DetermineTargets = targets
End Function

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

Function GenerateFileName(analystInitials As String, targets As String) As String
    ' Generate filename following convention: yyyymmdd-analyst_initials-targets.xlsx
    Dim dateStr As String
    Dim fileName As String
    
    dateStr = Format(Date, "yyyymmdd")
    fileName = dateStr & "-" & analystInitials & "-" & targets & ".xlsx"
    
    GenerateFileName = fileName
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
    
    proceed = True
    
    ' Get analyst name first
    analystName = InputBox("Enter analyst name (for file naming):", "Analyst Information", Environ("USERNAME"))
    If analystName = "" Then
        proceed = False
        GoTo ExitFunction
    End If
    
    ' Show initial message about parameters
    If MsgBox("Do you want to use default processing parameters?" & vbCrLf & vbCrLf & _
              "Defaults:" & vbCrLf & _
              "Initial Sample Volume: 40 ml" & vbCrLf & _
              "Concentrated Volume: 15 ml" & vbCrLf & _
              "Process Control Volume: 0.1 ml" & vbCrLf & _
              "RNA Elution Volume: 100 �l" & vbCrLf & _
              "qPCR Template Volume: 4 �l", vbYesNo + vbQuestion, "Processing Parameters") = vbYes Then
        
        ' Use defaults
        initialSampleVol = 40
        concentratedVol = 15
        processControlVol = 0.1
        extractionElutionVol = 100
        qpcrTemplateVol = 4
        
    Else
        ' Get custom parameters
        response = InputBox("Enter Initial Sample Volume (ml):", "Processing Parameters", "40")
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        initialSampleVol = CDbl(response)
        
        response = InputBox("Enter Concentrated Volume (ml):", "Processing Parameters", "15")
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        concentratedVol = CDbl(response)
        
        response = InputBox("Enter Process Control Volume (ml):", "Processing Parameters", "0.1")
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        processControlVol = CDbl(response)
        
        response = InputBox("Enter RNA Elution Volume (�l):", "Processing Parameters", "100")
        If response = "" Then proceed = False: GoTo ExitFunction
        If Not IsNumeric(response) Then proceed = False: GoTo ExitFunction
        extractionElutionVol = CDbl(response)
        
        response = InputBox("Enter qPCR Template Volume (�l):", "Processing Parameters", "4")
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
    resultsWs.Cells(currentRow, 1).Value = "N1 qPCR Analysis Results"
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

Sub ProcessSamplesWithQC(ws As Worksheet, resultsWs As Worksheet, dataStartRow As Long, lastRow As Long, _
                        colSampleName As Long, colTargetName As Long, colTask As Long, colCqMean As Long, _
                        colQuantityMean As Long, colQuantityStd As Long, _
                        n1Slope As Double, n1RSquared As Double, n1Intercept As Double, n1CurveValid As Boolean, _
                        mengoSlope As Double, mengoRSquared As Double, mengoIntercept As Double, mengoCurveValid As Boolean, _
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
    Dim sampleData As Collection
    Set uniqueSamples = New Collection
    Set sampleData = New Collection
    
    ' First pass: collect unique samples and their data
    For i = dataStartRow + 1 To lastRow
        sampleName = ws.Cells(i, colSampleName).Value
        targetName = UCase(ws.Cells(i, colTargetName).Value)
        task = UCase(ws.Cells(i, colTask).Value)
        
        ' Skip empty rows, standards, and controls
        If sampleName <> "" And Not IsControl(sampleName, task) And targetName = "N1" Then
            baseSampleName = GetBaseSampleName(sampleName)
            dilutionFactor = GetDilutionFactor(sampleName)
            
            ' Try to add to unique samples collection
            On Error Resume Next
            uniqueSamples.Add baseSampleName, baseSampleName
            On Error GoTo 0
            
            ' Store sample data
            Dim sampleInfo As String
            sampleInfo = baseSampleName & "|" & CStr(dilutionFactor) & "|" & sampleName
            sampleData.Add sampleInfo
        End If
    Next i
    
    ' Add standard curve QC section
    currentRow = resultsWs.Cells(resultsWs.Rows.Count, 1).End(xlUp).Row + 2
    resultsWs.Cells(currentRow, 1).Value = "Standard Curve Quality Control:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    ' N1 Standard Curve
    resultsWs.Cells(currentRow, 1).Value = "N1 Standard Curve:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Slope:"
    resultsWs.Cells(currentRow, 2).Value = Format(n1Slope, "0.000")
    resultsWs.Cells(currentRow, 3).Value = "Required: " & SLOPE_MIN & " to " & SLOPE_MAX
    resultsWs.Cells(currentRow, 4).Value = IIf(n1Slope >= SLOPE_MIN And n1Slope <= SLOPE_MAX, "PASS", "FAIL")
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "R�:"
    resultsWs.Cells(currentRow, 2).Value = Format(n1RSquared, "0.000")
    resultsWs.Cells(currentRow, 3).Value = "Required: =" & R_SQUARED_MIN
    resultsWs.Cells(currentRow, 4).Value = IIf(n1RSquared >= R_SQUARED_MIN, "PASS", "FAIL")
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Overall Status:"
    resultsWs.Cells(currentRow, 2).Value = IIf(n1CurveValid, "VALID", "INVALID")
    currentRow = currentRow + 2
    
    ' Mengo Standard Curve
    resultsWs.Cells(currentRow, 1).Value = "Mengo Standard Curve:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Slope:"
    resultsWs.Cells(currentRow, 2).Value = Format(mengoSlope, "0.000")
    resultsWs.Cells(currentRow, 3).Value = "Required: " & SLOPE_MIN & " to " & SLOPE_MAX
    resultsWs.Cells(currentRow, 4).Value = IIf(mengoSlope >= SLOPE_MIN And mengoSlope <= SLOPE_MAX, "PASS", "FAIL")
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "R�:"
    resultsWs.Cells(currentRow, 2).Value = Format(mengoRSquared, "0.000")
    resultsWs.Cells(currentRow, 3).Value = "Required: =" & R_SQUARED_MIN
    resultsWs.Cells(currentRow, 4).Value = IIf(mengoRSquared >= R_SQUARED_MIN, "PASS", "FAIL")
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Overall Status:"
    resultsWs.Cells(currentRow, 2).Value = IIf(mengoCurveValid, "VALID", "INVALID")
    currentRow = currentRow + 2
    
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
    
    resultsWs.Cells(currentRow, 1).Value = "RNA Elution Volume (�l):"
    resultsWs.Cells(currentRow, 2).Value = extractionElutionVolume
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "qPCR Template Volume (�l):"
    resultsWs.Cells(currentRow, 2).Value = qpcrTemplateVolume
    currentRow = currentRow + 1
    
    resultsWs.Cells(currentRow, 1).Value = "Base Concentration Factor:"
    resultsWs.Cells(currentRow, 2).Value = Format(totalConcentrationFactor, "0.0")
    currentRow = currentRow + 3
    
    ' Create results table header
    resultsWs.Cells(currentRow, 1).Value = "Sample Results:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    
    ' Table headers
    resultsWs.Cells(currentRow, 1).Value = "Sample Name"
    resultsWs.Cells(currentRow, 2).Value = "Dilution Factor"
    resultsWs.Cells(currentRow, 3).Value = "Cq Mean"
    resultsWs.Cells(currentRow, 4).Value = "Quantity Mean"
    resultsWs.Cells(currentRow, 5).Value = "Final Conc. (copies/L)"
    resultsWs.Cells(currentRow, 6).Value = "RT-PCR Inhibition (%)"
    resultsWs.Cells(currentRow, 7).Value = "Mengo Recovery (%)"
    resultsWs.Cells(currentRow, 8).Value = "QC Status"
    resultsWs.Cells(currentRow, 9).Value = "Calculation Details"
    
    ' Make headers bold
    For i = 1 To 9
        resultsWs.Cells(currentRow, i).Font.Bold = True
    Next i
    currentRow = currentRow + 1
    
    ' Process each unique sample
    Dim j As Integer
    For j = 1 To uniqueSamples.Count
        baseSampleName = uniqueSamples(j)
        
        ' Find the best dilution for this sample (lowest dilution with valid data)
        Dim bestSampleName As String
        Dim bestDilution As Double
        Dim bestQuantity As Double
        Dim bestCq As Double
        
        bestDilution = 999999 ' Start with very high number
        
        ' Look through all variants of this sample
        For i = dataStartRow + 1 To lastRow
            sampleName = ws.Cells(i, colSampleName).Value
            targetName = UCase(ws.Cells(i, colTargetName).Value)
            
            If GetBaseSampleName(sampleName) = baseSampleName And targetName = "N1" Then
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
        
        ' Calculate concentrations and QC
        If bestSampleName <> "" Then
            Dim finalConcentration As Double
            finalConcentration = bestQuantity * totalConcentrationFactor * bestDilution
            
            ' Calculate RT-PCR inhibition and Mengo recovery
            Dim inhibition As Double, recovery As Double
            Dim inhibitionPercent As String, recoveryPercent As String
            Dim inhibitionValid As Boolean, recoveryValid As Boolean
            Dim qcStatus As String
            
            ' RT-PCR Inhibition calculation
            Dim sampleEcCq As Double, controlEcCq As Double, deltaCq As Double
            sampleEcCq = GetSampleEcRnaCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean, baseSampleName)
            controlEcCq = GetControlEcRnaCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean)
            
            If sampleEcCq > 0 And controlEcCq > 0 Then
                deltaCq = sampleEcCq - controlEcCq
                inhibition = (1 - (10 ^ (deltaCq / n1Slope))) * 100
                inhibitionPercent = Format(inhibition, "0.1") & "%"
                inhibitionValid = (inhibition < INHIBITION_THRESHOLD)
            Else
                inhibitionPercent = "Not Calculated"
                inhibitionValid = False
            End If
            
            ' Mengo Recovery calculation
            Dim sampleMengoCq As Double, processControlMengoCq As Double
            sampleMengoCq = GetSampleMengoCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean, baseSampleName)
            processControlMengoCq = GetProcessControlMengoCq(ws, dataStartRow, lastRow, colSampleName, colTargetName, colCqMean)
            
            If sampleMengoCq > 0 And processControlMengoCq > 0 Then
                deltaCq = sampleMengoCq - processControlMengoCq
                Dim samplePathDilution As Double, expectedDilutionDifference As Double
                samplePathDilution = concentratedVolume / processControlVolume
                expectedDilutionDifference = samplePathDilution / 1 ' Control is direct extraction
                
                Dim rawRecovery As Double
                rawRecovery = (10 ^ (deltaCq / mengoSlope)) * 100
                recovery = rawRecovery * expectedDilutionDifference
                recoveryPercent = Format(recovery, "0.1") & "%"
                recoveryValid = (recovery > RECOVERY_THRESHOLD)
            Else
                recoveryPercent = "Not Calculated"
                recoveryValid = False
            End If
            
            ' Determine QC status
            If n1CurveValid And mengoCurveValid And ntcIssues = "" Then
                If inhibitionPercent <> "Not Calculated" And recoveryPercent <> "Not Calculated" Then
                    If inhibitionValid And recoveryValid Then
                        qcStatus = "Valid"
                    Else
                        qcStatus = "QC Fail"
                    End If
                Else
                    qcStatus = "Incomplete QC"
                End If
            Else
                qcStatus = "Invalid"
            End If
            
            ' Calculate details
            Dim calcDetails As String
            calcDetails = Format(bestQuantity, "#,##0") & " � " & Format(totalConcentrationFactor, "0.0") & " � " & bestDilution & " = " & Format(finalConcentration, "#,##0")
            
            ' Write results
            resultsWs.Cells(currentRow, 1).Value = baseSampleName
            resultsWs.Cells(currentRow, 2).Value = IIf(bestDilution = 1, "None", "1:" & bestDilution)
            resultsWs.Cells(currentRow, 3).Value = IIf(bestCq > 0, Format(bestCq, "0.000"), "N/A")
            resultsWs.Cells(currentRow, 4).Value = Format(bestQuantity, "#,##0")
            resultsWs.Cells(currentRow, 5).Value = Format(finalConcentration, "#,##0")
            resultsWs.Cells(currentRow, 6).Value = inhibitionPercent
            resultsWs.Cells(currentRow, 7).Value = recoveryPercent
            resultsWs.Cells(currentRow, 8).Value = qcStatus
            resultsWs.Cells(currentRow, 9).Value = calcDetails
            
            currentRow = currentRow + 1
        End If
    Next j
    
    ' Add notes
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "QC Criteria:"
    resultsWs.Cells(currentRow, 1).Font.Bold = True
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "� RT-qPCR Inhibition: <" & INHIBITION_THRESHOLD & "% acceptable"
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "� Mengo Recovery: >" & RECOVERY_THRESHOLD & "% acceptable"
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "� Standard Curves: Slope " & SLOPE_MIN & " to " & SLOPE_MAX & ", R� =" & R_SQUARED_MIN
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "� NTC: No amplification detected"
    currentRow = currentRow + 1
    resultsWs.Cells(currentRow, 1).Value = "* Valid with QC warnings indicates passing standard curves and NTC but failing inhibition/recovery"
    resultsWs.Cells(currentRow, 1).Font.Italic = True
    
End Sub

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
        
        If GetBaseSampleName(sampleName) = baseSampleName And targetName = "EC-N1" Then
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
        
        If (InStr(sampleName, "WATER") > 0 Or InStr(sampleName, "CONTROL") > 0) And targetName = "EC-N1" Then
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
        If InStr(resultsWs.Cells(i, 1).Value, "Sample Results:") > 0 Then
            tableStartRow = i + 1 ' Header row
            Exit For
        End If
    Next i
    
    If tableStartRow > 0 Then
        ' Find end of table
        tableEndRow = resultsWs.Cells(resultsWs.Rows.Count, 1).End(xlUp).Row
        
        ' Apply table formatting
        With resultsWs.Range(resultsWs.Cells(tableStartRow, 1), resultsWs.Cells(tableEndRow, 9))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
        ' Header row formatting
        With resultsWs.Range(resultsWs.Cells(tableStartRow, 1), resultsWs.Cells(tableStartRow, 9))
            .Interior.Color = RGB(200, 200, 200)
            .Font.Bold = True
        End With
        
        ' Color code QC status column
        For i = tableStartRow + 1 To tableEndRow
            If resultsWs.Cells(i, 8).Value = "Valid" Or InStr(resultsWs.Cells(i, 8).Value, "Valid") > 0 Then
                resultsWs.Cells(i, 8).Interior.Color = RGB(200, 255, 200) ' Light green
            ElseIf resultsWs.Cells(i, 8).Value = "Invalid" Then
                resultsWs.Cells(i, 8).Interior.Color = RGB(255, 200, 200) ' Light red
            End If
        Next i
    End If
    
    ' Set column widths
    resultsWs.Columns("A:A").ColumnWidth = 18 ' Sample Name
    resultsWs.Columns("B:B").ColumnWidth = 12 ' Dilution Factor
    resultsWs.Columns("C:C").ColumnWidth = 10 ' Cq Mean
    resultsWs.Columns("D:D").ColumnWidth = 12 ' Quantity Mean
    resultsWs.Columns("E:E").ColumnWidth = 15 ' Final Concentration
    resultsWs.Columns("F:F").ColumnWidth = 15 ' Inhibition
    resultsWs.Columns("G:G").ColumnWidth = 15 ' Recovery
    resultsWs.Columns("H:H").ColumnWidth = 12 ' QC Status
    resultsWs.Columns("I:I").ColumnWidth = 30 ' Calculation Details
    
End Sub




