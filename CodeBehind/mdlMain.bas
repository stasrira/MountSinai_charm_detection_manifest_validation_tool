Attribute VB_Name = "mdlMain"
Option Explicit

Public Const Version = "1.04"

Private Const ExistingDetectionsWrkSh = "COVID_Detection_Existing"
Private Const DetectionFileWrkSh = "Detection_File"
Private Const TempLoadWrkSh = "temp_load"
Private Const ConfigWrkSheet = "config"
Private Const DictionaryWrkSheet = "dictionary"

Public Global_Validated As Boolean

Function SameTimepointExists(existing_timepoints As String, timepoint As String) As String
    Dim arr As Variant
    
    arr = Split(existing_timepoints, ",")
    If ArrLength(arr) > 0 Then
        SameTimepointExists = CStr(IsValueInArray(arr, timepoint))
    Else
        SameTimepointExists = CStr(False)
    End If
    
End Function

Function IsValueInArray(arr As Variant, value As Variant)
    If IsNumeric(Application.Match(value, arr, 0)) Then
        IsValueInArray = True
    Else
        IsValueInArray = False
    End If
End Function

Function IsNotTimepointLatest(existing_timepoints As String, timepoint As String, Optional NumLeadCharsToRemove As Integer = 0) As String
    Dim arr As Variant
    Dim lastElement As String
    Dim lastElementNum As Long, timepointNum As Long
    Dim lastElementReady As Boolean, timepointReady As Boolean

    
    arr = Split(existing_timepoints, ",")
    
    If ArrLength(arr) > 0 Then
    
        'assumption that all timepoints have same structure - having the same (important) prefix letter(s) and the rest is a numeric value
        lastElement = arr(ArrLength(arr) - 1)
        
        If Len(lastElement) > NumLeadCharsToRemove Then
            lastElement = Right(lastElement, Len(lastElement) - NumLeadCharsToRemove)
            If IsNumeric(lastElement) Then
                lastElementNum = CLng(lastElement)
                lastElementReady = True
            End If
        End If
        
        If Len(timepoint) > NumLeadCharsToRemove Then
            timepoint = Right(timepoint, Len(timepoint) - NumLeadCharsToRemove)
            If IsNumeric(timepoint) Then
                timepointNum = CLng(timepoint)
                timepointReady = True
            End If
        End If
        
        If lastElementReady And timepointReady Then
            IsNotTimepointLatest = CStr(Not timepointNum >= lastElementNum)
        Else
            IsNotTimepointLatest = "N/A"
        End If
    Else
        IsNotTimepointLatest = CStr(Not True)
    End If
    
End Function

Function ExistingTimepointsForSubject(subject_id As String) As String
    Dim exist_tps() As String
    Dim rng As Range
    Dim subject_id_col As String
    Dim timepoint_offset_num As String
    
    subject_id_col = GetConfigParameterValueB("ExistingDetection_SubjectID_Column")
    timepoint_offset_num = GetConfigParameterValueB("ExistingDetection_TimepointColumn_Offset_From_Subject")
    
    Set rng = Worksheets(ExistingDetectionsWrkSh).Range(subject_id_col & ":" & subject_id_col) 'COVID_Detection_Existing '"F:F"
    
    exist_tps = FindAllValuesOrLocationRows(rng, Trim(subject_id), "0," & timepoint_offset_num, True) '"0,-4"
    
    ExistingTimepointsForSubject = Join(exist_tps, ",")
    
End Function

Function NewTimepointsForSubjectInManifest(subject_id As String) As String
    Dim exist_tps() As String
    Dim rng As Range
    Dim subject_id_col As String
    Dim timepoint_offset_num As String
    Dim rows_num As Integer
    
    If Len(Trim(subject_id)) = 0 Then
        'do not proceed if the subject id is blank
        NewTimepointsForSubjectInManifest = ""
        Exit Function
    End If
    
    rows_num = DetectionFileLoadedRows()
    
    subject_id_col = GetConfigParameterValueB("DetectionFile_SubjectID_Column")
    timepoint_offset_num = GetConfigParameterValueB("DetectionFile_TimepointColumn_Offset_From_Subject")
    
    Set rng = Worksheets(DetectionFileWrkSh).Range(subject_id_col & "1:" & subject_id_col & CStr(rows_num)) 'COVID_Detection_Existing '"A:A"
    
    exist_tps = FindAllValuesOrLocationRows(rng, Trim(subject_id), "0," & timepoint_offset_num, True, False) '"0,1"
    
    NewTimepointsForSubjectInManifest = Join(exist_tps, ",")
    
End Function

Function DuplicatedEntriesInManifest(subject_id As String) As String
    Dim tps() As String, tps_sorted() As String
    Dim row_nums() As String, dupl_rows() As String
    Dim rng As Range
    Dim i As Integer, j As Integer
    Dim duplicates() As String, duplicates_report() As String
    Dim aInitialised As Boolean, aInitialised2 As Boolean
    Dim subject_id_col As String
    Dim timepoint_offset_num As String
    Dim rows_num As Integer
        
    If Len(Trim(subject_id)) = 0 Then
        'do not proceed if the subject id is blank
        DuplicatedEntriesInManifest = ""
        Exit Function
    End If
    
    rows_num = DetectionFileLoadedRows()
    
    subject_id_col = GetConfigParameterValueB("DetectionFile_SubjectID_Column")
    timepoint_offset_num = GetConfigParameterValueB("DetectionFile_TimepointColumn_Offset_From_Subject")
    
    Set rng = Worksheets(DetectionFileWrkSh).Range(subject_id_col & "1:" & subject_id_col & CStr(rows_num)) 'COVID_Detection_Existing '"A:A"
    
    tps = FindAllValuesOrLocationRows(rng, Trim(subject_id), "0," & timepoint_offset_num, False, False) '"0,1"
    row_nums = FindAllValuesOrLocationRows(rng, Trim(subject_id), "0," & timepoint_offset_num, False, True) '"0,1"
    
    tps_sorted = tps
    QuickSort tps_sorted, 0, UBound(tps_sorted) 'sort the values of the array
    
    If ArrLength(tps_sorted) > 1 Then
        'since array is sorted, duplicated values will be next to each other.
        For i = 1 To ArrLength(tps_sorted) - 1
            If tps_sorted(i - 1) = tps_sorted(i) Then
                'duplicate is found
                If Not IsInArray(CStr(tps_sorted(i)), duplicates) Then 'check if item no in array yet
                    'Redim array size
                    If Not aInitialised Then
                        ReDim Preserve duplicates(0)
                        ReDim Preserve duplicates_report(0)
                        aInitialised = True
                    Else
                        ReDim Preserve duplicates(ArrLength(duplicates))
                        ReDim Preserve duplicates(ArrLength(duplicates_report))
                    End If
                    
                    'loop through all found timepoint values and identify rows the duplicates values are located on
                    For j = 0 To ArrLength(tps) - 1
                        If tps(j) = tps_sorted(i) Then
                            'Redim array size
                            If Not aInitialised2 Then
                                ReDim Preserve dupl_rows(0)
                                aInitialised2 = True
                            Else
                                ReDim Preserve dupl_rows(ArrLength(dupl_rows))
                            End If
                            dupl_rows(ArrLength(dupl_rows) - 1) = row_nums(j)
                        End If
                    Next
                    aInitialised2 = False 'reset this flag to False
                    
                    duplicates(ArrLength(duplicates) - 1) = tps_sorted(i)
                    duplicates_report(ArrLength(duplicates_report) - 1) = tps_sorted(i) & "(row#:" & Join(dupl_rows, ",") & ")"
                End If
            End If
        Next
    End If
    
    DuplicatedEntriesInManifest = Join(duplicates_report, ",")
    
End Function

Function FindAllValuesOrLocationRows(rng As Range, _
                        What As Variant, _
                        Optional offset_row_col As String = "0,0", _
                        Optional sort_output As Boolean = False, _
                        Optional return_row_number As Boolean = False, _
                        Optional LookIn As XlFindLookIn = xlValues, _
                        Optional LookAt As XlLookAt = xlWhole, _
                        Optional SearchOrder As XlSearchOrder = xlByRows, _
                        Optional SearchDirection As XlSearchDirection = xlNext, _
                        Optional MatchCase As Boolean = False, _
                        Optional MatchByte As Boolean = False, _
                        Optional SearchFormat As Boolean = False) As String()
                        
                        'xlByColumns
                        
    Dim SearchResult As Range
    Dim firstMatch As String
    Dim out_arr() As String
    Dim aInitialised As Boolean
    Dim offSetcol As Integer, offSetrow As Integer
    Dim curCell As Range
    
    offSetrow = Split(offset_row_col, ",")(0)
    offSetcol = Split(offset_row_col, ",")(1)
            
    With rng
        Set SearchResult = .Find(What, rng.Cells(1, 1), LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
        If Not SearchResult Is Nothing Then
            firstMatch = SearchResult.Address
            'ReDim Preserve out_arr(0)
            
            Do
                If Not aInitialised Then
                    ReDim Preserve out_arr(0)
                    aInitialised = True
                Else
                    ReDim Preserve out_arr(ArrLength(out_arr))
                End If
                
                If Not return_row_number Then
                    out_arr(ArrLength(out_arr) - 1) = SearchResult.offset(offSetrow, offSetcol).value
                Else
                    out_arr(ArrLength(out_arr) - 1) = SearchResult.Row
                End If
                
                'Set SearchResult = .FindNext(SearchResult) 'this approach did not work
                Set SearchResult = .Find(What, SearchResult, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
                If Not SearchResult Is Nothing Then
                    If SearchResult.Address = firstMatch Then
                        Set SearchResult = Nothing
                    End If
                End If
            Loop While Not SearchResult Is Nothing 'And SearchResult.Address <> firstMatch
        Else
            'return blank array if there was no match
            ReDim Preserve out_arr(0)
        End If
    End With
    
    If sort_output Then
        QuickSort out_arr, 0, UBound(out_arr) 'sort the values of the array
    End If
    
    FindAllValuesOrLocationRows = out_arr
End Function

Private Function ArrLength(a As Variant) As Integer
   If IsEmpty(a) Then
      ArrLength = 0
   Else
      ArrLength = UBound(a) - LBound(a) + 1
   End If
End Function

Private Sub QuickSort(ByRef Field() As String, ByVal LB As Long, ByVal UB As Long)
    Dim P1 As Long, P2 As Long, Ref As String, TEMP As String

    P1 = LB
    P2 = UB
    Ref = Field((P1 + P2) / 2)

    Do
        Do While (Field(P1) < Ref)
            P1 = P1 + 1
        Loop

        Do While (Field(P2) > Ref)
            P2 = P2 - 1
        Loop

        If P1 <= P2 Then
            TEMP = Field(P1)
            Field(P1) = Field(P2)
            Field(P2) = TEMP

            P1 = P1 + 1
            P2 = P2 - 1
        End If
    Loop Until (P1 > P2)

    If LB < P2 Then Call QuickSort(Field, LB, P2)
    If P1 < UB Then Call QuickSort(Field, P1, UB)
End Sub

Function checkConditionalFormat(sheet As String, column_letter As String, condition_val As Boolean) As Boolean
    Dim ws As Worksheet
    Dim col As Range
    
    Set ws = Worksheets(sheet)
    Set col = ws.Range(column_letter & ":" & column_letter)
    
    checkConditionalFormat = IsNumeric(Application.Match(condition_val, col, 0))
    
End Function

Function checkConditionalFormatCount(sheet As String, column_letter As String, condition_val As Boolean) As Integer
    Dim ws As Worksheet
    Dim col As Range
    
    Set ws = Worksheets(sheet)
    Set col = ws.Range(column_letter & ":" & column_letter)
    checkConditionalFormatCount = Application.WorksheetFunction.CountIf(col, condition_val)
    'checkConditionalFormatCount = Application.WorksheetFunction.CountIf(col, condition_val & "*")
    
End Function

Private Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    On Error GoTo errLab
    IsInArray = Not IsError(Application.Match(stringToBeFound, arr, 0))
    Exit Function
errLab:
    IsInArray = False
End Function

Private Function GetConfigParameterValueB(cfg_param_name As String, Optional wb As Workbook = Nothing) As String
    'retrieve value from column B of the config tab
    GetConfigParameterValueB = GetConfigParameterValueByColumn(cfg_param_name, "B", wb)
End Function

Private Function GetConfigParameterValueC(cfg_param_name As String, Optional wb As Workbook = Nothing) As String
    'retrieve value from column C of the config tab
    GetConfigParameterValueC = GetConfigParameterValueByColumn(cfg_param_name, "C", wb)
End Function

Private Function GetConfigParameterValueByColumn(cfg_param_name As String, column_letter As String, Optional wb As Workbook = Nothing) As String
    Dim cfg_row As Integer
    Dim out_val As String
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(ConfigWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(ConfigWrkSheet)
    End If
    
    'get config value to identify column letter on the shipment tab
    cfg_row = FindRowNumberOfConfigParam(cfg_param_name, wb)
    If cfg_row > 0 Then
        'get configuration value
        out_val = ws_cfg.Range(column_letter & CStr(cfg_row))
    Else
        out_val = ""
    End If
    GetConfigParameterValueByColumn = out_val
End Function

'searches for a given parameter name on the config page and returns the row number it was found on
Private Function FindRowNumberOfConfigParam(param_name As String, Optional wb As Workbook = Nothing) As Integer
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(ConfigWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(ConfigWrkSheet)
    End If
    
    If IsError(Application.Match(param_name, ws_cfg.Range("A:A"), 0)) Then
        FindRowNumberOfConfigParam = 0
    Else
        FindRowNumberOfConfigParam = Application.Match(param_name, ws_cfg.Range("A:A"), 0)
    End If
    
End Function

Public Sub ImportDetectionFile()
    Dim iResponse As Integer
    Dim importFileOutcome As Boolean
    Dim strFileToOpen As String
    Dim tStart As Date, tEnd As Date
    
    'confirm if user want to proceed.
    iResponse = MsgBox("The system is about to start importing COVID Detection Manifest file to the 'Detection_File' tab. " _
                & vbCrLf & "- Any currently existing COVID Detection data will be overwritten wih the new data!" _
                & vbCrLf & "- The current 'Date Received' date will be set to: " & GetDateReceived() & ". If it is not desired, cancel this operation and update the 'Date_received_value' field on the 'Config' tab." _
                & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'." & vbCrLf & vbCrLf _
                & "Note: " _
                & vbCrLf & "- This process might take prolonged time (up to 1 or 2 minutes), depending on the number of rows being imported. " _
                & vbCrLf & "- The automatic verification of all validation rules will be run immediately after importing. " _
                & vbCrLf & "- A final confirmation will be displayed upon completion." _
                & vbCrLf & "- Some screen flickering might occur during the process. ", _
                vbOKCancel + vbInformation, "CHARM COVID Detection Validation")
    
    If iResponse <> vbOK Then
        'exit sub based on user's response
        Exit Sub
    End If
    
    'select a file to be loaded
    strFileToOpen = Application.GetOpenFilename _
        (Title:="Please choose a COVID Detectio file to open", _
        FileFilter:="Excel Files *.xlsx* (*.xlsx*), Excel 2003 Files *.xls* (*.xls*),")
    
    If strFileToOpen = "False" Then
        Exit Sub
    End If
    
    tStart = Now()
    'Debug.Print (tStart)
    
    importFileOutcome = ImportFile(strFileToOpen, Worksheets(DetectionFileWrkSh))
    
    If importFileOutcome Then
        'proceed here only if the file was successfully loaded
        Worksheets(DetectionFileWrkSh).Activate 'bring focus to the "logs" tab
        Worksheets(DetectionFileWrkSh).Cells(1, 1).Activate 'bring focus to the first cell on the sheet
        
        'MsgBox "File loading was completed. Proceeding to refreshing all validation results.", vbInformation, "CHARM COVID Detection Validation"
        
        RefreshWorkbookData True 'refresh all validation rules (avoiding showing a confirmation upon completion).
        
        tEnd = Now()
        'Debug.Print (tEnd)
        
        MsgBox "File loading was fully completed and all validation rules were refreshed." _
            & vbCrLf & "Execution time: " & getTimeLength(tStart, tEnd) _
            , vbInformation, "CHARM COVID Detection Validation"
    End If
    
End Sub

Private Function ImportFile(strFileToOpen As String, ws_target As Worksheet) As Boolean ', file_type_to_open As String
    'Dim strFileToOpen As String
    
    On Error GoTo ErrHandler 'commented to test, need to be uncommented
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim s As Worksheet
    
    Set s = Worksheets(TempLoadWrkSh)
    
'    'select a file to be loaded
'    strFileToOpen = Application.GetOpenFilename _
'        (Title:="Please choose a " & file_type_to_open & " file to open", _
'        FileFilter:="Excel Files *.xlsx* (*.xlsx*), Excel 2003 Files *.xls* (*.xls*),")
'
'    If strFileToOpen = "False" Then
'        GoTo ExitMark
'    End If
    
    s.Cells.Clear 'delete everything on the target worksheet
    
    CopyDataFromFile s, strFileToOpen 'copy date of the main sheet from the source file to the temp_load sheet
    
    DeleteBlankRows s 'clean blank rows of just imported file on the temp_load sheet
    
    CopySelectedColumnToTargetSheet s, Worksheets(DetectionFileWrkSh), GetConfigParameterValueB("Participant_ID_new_file_mapping") 'copy Pparticipant Id column '"A:A"
    CopySelectedColumnToTargetSheet s, Worksheets(DetectionFileWrkSh), GetConfigParameterValueB("Timepoint_new_file_mapping") 'copy Pparticipant Id column '"B:B"
    CopySelectedColumnToTargetSheet s, Worksheets(DetectionFileWrkSh), GetConfigParameterValueB("Results_new_file_mapping") 'copy Pparticipant Id column '"C:C"
    FillSelectedColumnWithConstantValue GetDateReceived(), Worksheets(DetectionFileWrkSh), GetConfigParameterValueB("Date_received_column"), s.UsedRange.Rows.Count
    FillSelectedColumnWithConstantValue GetFileNameFromPath(strFileToOpen), Worksheets(DetectionFileWrkSh), GetConfigParameterValueB("File_Name_Column"), s.UsedRange.Rows.Count
    
'    RefreshWorkbookData

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
'    MsgBox "CHARM " & file_type_to_open & " file " & vbCrLf _
'            & strFileToOpen & vbCrLf _
'            & " was successfully loaded to the '" & ws_target.Name & "' tab." & vbCrLf & vbCrLf _
'            & "Note: you might need to review settings of the 'config' tab to make sure that those are set correctly. " _
'            & "Please pay special attention to the highlighed rows."
    
    ImportFile = True
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbCritical
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ImportFile = False
End Function

'This sub opens specified file and loads it contents to a specified worksheet
Private Sub CopyDataFromFile(ws_target As Worksheet, _
                    src_file_path As String, _
                    Optional src_worksheet_name As String = "")
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim src As Workbook
    Dim path As String
    
    ' OPEN THE SOURCE EXCEL WORKBOOK IN "READ ONLY MODE".
    Set src = Workbooks.Open(src_file_path, True, True)
    
    If src_worksheet_name = "" Then
        src_worksheet_name = src.Worksheets(1).Name
    End If
    
    src.Worksheets(src_worksheet_name).Cells.Copy 'copy into a clipboard
    ws_target.Cells.PasteSpecial Paste:=xlPasteAll 'paste to the worksheet
    Application.CutCopyMode = False 'clean clipboard
    
  
    ' CLOSE THE SOURCE FILE.
    src.Close False             ' FALSE - DON'T SAVE THE SOURCE FILE.
    Set src = Nothing
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbCritical
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub CopySelectedColumnToTargetSheet(source As Worksheet, Target As Worksheet, mapping As String)
    Dim copy_cols() As String
    Dim src_col As Range, dst_col As Range
    Dim src_used_rows As Integer, dest_used_rows As Integer
    
    src_used_rows = source.UsedRange.Rows.Count
    dest_used_rows = Target.UsedRange.Rows.Count
    
    copy_cols = Split(mapping, ":")
    If ArrLength(copy_cols) > 1 Then
        Set src_col = source.Range(copy_cols(0) & "2:" & copy_cols(0) & CStr(src_used_rows))
        Set dst_col = Target.Range(copy_cols(1) & "2:" & copy_cols(1) & CStr(dest_used_rows))
    End If
    dst_col.Clear
    
    src_col.Cells.Copy 'copy into a clipboard
    dst_col.Cells.PasteSpecial Paste:=xlPasteValues 'paste to the worksheet
    
End Sub

Private Sub FillSelectedColumnWithConstantValue(update_value As String, Target As Worksheet, column_to_fill As String, fill_rows_num As Integer)
    Dim dst_col As Range
    Dim dest_used_rows As Integer
    
    'clean target column first
    dest_used_rows = Target.UsedRange.Rows.Count
    Set dst_col = Target.Range(column_to_fill & "2:" & column_to_fill & CStr(dest_used_rows))
    dst_col.Clear
    
    'update destination columns with the given value
    Set dst_col = Target.Range(column_to_fill & "2:" & column_to_fill & CStr(fill_rows_num))
    dst_col.value = update_value
    
End Sub

Public Sub RefreshWorkbookData(Optional hideConfirmation As Boolean = False)
    Dim refresh_db As String
    Dim tStart As Date, tEnd As Date
    Dim bPopulateValidColumns As Boolean
    Dim sMsg As String, iMsg As Integer
    
    tStart = Now()
        
    'clear out any applied filters
    ResetFilters Worksheets(DetectionFileWrkSh)
    
    'apply dictionary conversion to the values of the Detection column
    VerifyDetectionValues
    
    refresh_db = GetConfigParameterValueB("Run Database refresh link on a fly")
    If UCase(refresh_db) = "TRUE" Then
        RefreshDBConnections True
    End If
    
    'populate main validation columns
    bPopulateValidColumns = PopulateValidationColumns
    
    'recalculate whole workbook to make sure manifest and metadata sheets are filled properly
    Application.CalculateFullRebuild
    
    Worksheets(DetectionFileWrkSh).Activate 'bring focus to the "logs" tab
    Worksheets(DetectionFileWrkSh).Cells(1, 1).Activate 'bring focus to the first cell on the sheet
    
    tEnd = Now()
    
    If Not hideConfirmation Then
        If bPopulateValidColumns Then
            sMsg = "All validation rules were refereshed. Check results on the 'Detection_File' tab." _
                    & vbCrLf & "Execution time: " & getTimeLength(tStart, tEnd)
            iMsg = vbInformation
        Else
            sMsg = "Some errors were reported during processing validation rules. Some of the validation rules might have not be verified." _
                    & vbCrLf & "Adjustments of the configuration parameters might resolve the problem - verify entries on the 'config' tab." _
                    & vbCrLf & "Once errors are resolved, re-run the validatoin process."
            iMsg = vbExclamation
        End If
        
'        MsgBox "All validation rules were refereshed. Check results on the 'Detection_File' tab." _
'            & vbCrLf & "Execution time: " & getTimeLength(tStart, tEnd) _
'            , vbExclamation, "CHARM COVID Detection Validation"
        MsgBox sMsg, iMsg, "CHARM COVID Detection Validation"
    End If
    
    Global_Validated = True
    
    'Application.ScreenUpdating = False
    'ActiveWorkbook.ForceFullCalculation
'    Application.ScreenUpdating = True
End Sub

Public Sub RefreshDBConnections(Optional hideConfirmation As Boolean = False)
    'refresh the database linked data
    ActiveWorkbook.RefreshAll
    
    If Not hideConfirmation Then
        MsgBox "All database linked data was refreshed. Check results on the 'COVID_Detection_Existing' tab.", vbInformation, "CHARM COVID Detection Validation"
    End If
End Sub

Private Function GetFileNameFromPath(path As String)
    Dim fso As New FileSystemObject
    Dim fileName As String
    
    fileName = fso.GetFileName(path)
    GetFileNameFromPath = fileName
End Function

Private Function GetDateReceived() As String
    Dim cfg_value As String
    
    cfg_value = GetConfigParameterValueB("Date_received_value")
    
    If UCase(cfg_value) = "AUTO" Then
        GetDateReceived = CStr(Date)
    Else
        If IsDate(cfg_value) Then
            GetDateReceived = cfg_value
        Else
            GetDateReceived = "N/D"
        End If
    End If
End Function

Public Function SavePreparedData() As dictionary
    'Dim abort As Boolean: abort = False
    Dim wb As Workbook, ws_source As Worksheet, ws_target As Worksheet
    Dim wb_source As Workbook
    Dim out_dict As New Scripting.dictionary
    Dim new_file_name As String, path As String
    Dim str1 As String
    Dim empty_file_flag As Boolean
    Dim msg_status As Integer
    Dim iResponse As Integer
    
    If Not Global_Validated Then
        'confirm if user want to proceed.
        iResponse = MsgBox("Some data interactoins were registered on the 'Data_Validation' tab. The displayed validation results might not be current. " _
                    & "It is suggested to re-run the 'Referesh Validatoin Results' process." _
                    & vbCrLf & vbCrLf & "Do you want to proceed with the Export procedure anyway? If not, click 'Cancel'.", _
                    vbOKCancel + vbInformation, "CHARM COVID Detection Validation")
        
        If iResponse <> vbOK Then
            'exit sub based on user's response
            Exit Function
        End If
    End If
    
    path = GetConfigParameterValueB("Save Created Metadata Files Path")
    'validate received path
    If Dir(path, vbDirectory) = "" Then
        str1 = "The path to the exporting directory:" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & "provided in the 'config' tab (see item named 'Save Created Metadata Files Path') cannot be reached. " _
                & "Please verify that path is valid and the current user has access to it." _
                & vbCrLf & vbCrLf & "Creation of the new detection export file was aborted!"
        MsgBox str1, vbCritical, "Export of Detection Data"
        
        Set SavePreparedData = out_dict
        Exit Function
    End If
    
    new_file_name = path & "\" & GetFileNameToSave()  'get the file name for the new excel file
    
    'confirm if user want to proceed.
    iResponse = MsgBox("The system is about to start exporting validated data from 'Detection_File' tab." & vbCrLf _
                & "Based on the tool's configuration settings, the output will be saved to the following file: " & vbCrLf & new_file_name _
                & vbCrLf & vbCrLf & "Do you want to proceed? If not, click 'Cancel'.", _
                vbOKCancel + vbInformation, "CHARM COVID Detection Validation")
                
    
    If iResponse <> vbOK Then
        'exit sub based on user's response
        Exit Function
    End If
    
    'On Error GoTo ErrHandler 'commented to test, need to be uncommented
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set wb_source = Application.ActiveWorkbook
    Set ws_source = Worksheets(DetectionFileWrkSh) 'set reference to the "metadata" sheet
    
    'AddLogEntry "Start creating new '" & file_type & "' file in '" & path & "' folder.", LogMsgType.info
    
    Set wb = CreateNewFile(new_file_name) ' create a new excel file and get a reference to it
    If Not wb Is Nothing Then
        Set ws_target = wb.Sheets(1) 'get reference to the worksheet of the new file
        
        ws_source.Range(GetConfigParameterValueB("DetectionFile_ColumnExportRange", wb_source)).Copy 'copy data from a source sheet to a memory
        ws_target.Cells.PasteSpecial Paste:=xlPasteValues 'paste data from memory to the target sheet as "values only"
        
        CleanCreatedFile ws_target, wb_source
        
        AddDummyColumnsToExport ws_target, wb_source
        
        wb.Save 'save the new file
        
        If ws_target.UsedRange.Rows.Count > 1 Then
            'if the cleaned worksheet contains at least on data row (beside the header), proceed here
            empty_file_flag = False 'set flag to save changes on closing
            
            str1 = "New detection export file was successfully created - " & wb.FullName '& ""
            msg_status = vbInformation
            out_dict.Add "msg", str1
            out_dict.Add "status", "OK" 'vbInformation
            
            'AddLogEntry str1, LogMsgType.info, wb_source
        Else
            empty_file_flag = True 'set flag not to save changes on closing
        
            str1 = "Newly created detection export file - " & wb.FullName & " - appears to be empty and will be deleted."
            msg_status = vbExclamation
            out_dict.Add "msg", str1
            out_dict.Add "status", "EMPTY" 'vbInformation
            
            'AddLogEntry str1, LogMsgType.warning, wb_source
        End If
        
        wb.Close 'close the new file
        
        MsgBox str1, msg_status, "Export of Detection Data"
        
        If empty_file_flag Then
            'delete just created empty file
            If DeleteFile(new_file_name) Then
                str1 = "Newly created detection export file - " & new_file_name & " - was successfully deleted."
                'AddLogEntry str1, LogMsgType.warning, wb_source
            Else
                str1 = "The application was not able to delete the newly created detection export file - " & wb.FullName & "."
                'AddLogEntry str1, LogMsgType.Error, wb_source
            End If
        End If
        
        Application.CutCopyMode = False 'clean clipboard
    Else
        str1 = "Existing file with the same name was present (" & new_file_name & "). " & vbCrLf & vbCrLf & "Creation of the new detection export file was skipped based on user's input."
        MsgBox str1, vbCritical, "Export of Detection Data"
        out_dict.Add "msg", str1
        out_dict.Add "status", "EXISTS" 'vbInformation
        'AddLogEntry str1, LogMsgType.warning
    End If
    'AddLogEntry "Finish creating new '" & file_type & "' file.", LogMsgType.info
    
    Set SavePreparedData = out_dict
    Exit Function
    
ErrHandler:
    MsgBox Err.Description, vbCritical
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

Private Sub CleanCreatedFile(ws_target As Worksheet, wb_source As Workbook)
    Dim abort As Boolean: abort = False
    Dim deleteColumns As String: deleteColumns = ""
    Dim formatDateColumns As String: formatDateColumns = ""
    Dim col As Variant, cfg_row As Integer
    
    deleteColumns = GetConfigParameterValueB("Delete Columns In Target", wb_source)
    formatDateColumns = GetConfigParameterValueB("Format Date Columns In Target", wb_source)
    
    If Len(Trim(deleteColumns)) > 0 Then
        'loop through columns and delete one by one
        For Each col In Split(deleteColumns, ",")
            col = Trim(col)
            If Len(col) > 0 Then
                ws_target.Columns(col).Delete
            End If
        Next
    End If
    
    If Len(Trim(formatDateColumns)) > 0 Then
        'loop through columns and delete one by one
        For Each col In Split(formatDateColumns, ",")
            col = Trim(col)
            If Len(col) > 0 Then
                ws_target.Columns(col).NumberFormat = "mm/dd/yyyy"
            End If
        Next
    End If
    
    DeleteBlankRows ws_target 'delete blank rows in the new file

End Sub

Private Sub DeleteBlankRows(ws_target As Worksheet)
    Dim SourceRange As Range
    Dim EntireRow As Range
    Dim i As Long, non_blanks As Long, empty_strings As Long
 
    Set SourceRange = ws_target.UsedRange ' Cells.End(xlToLeft)
 
    If Not (SourceRange Is Nothing) Then
        'Application.ScreenUpdating = False
 
        For i = SourceRange.Rows.Count To 1 Step -1
            Set EntireRow = SourceRange.Cells(i, 1).EntireRow
            non_blanks = Application.WorksheetFunction.CountA(EntireRow)
            empty_strings = Application.WorksheetFunction.CountIf(EntireRow, "")
            If non_blanks = 0 Or EntireRow.Cells.Count = empty_strings Then
                EntireRow.Delete
            'Else
                'Print ("Not blank row")
            End If
        Next
 
        'Application.ScreenUpdating = True
    End If
End Sub

Private Sub AddDummyColumnsToExport(ws_target As Worksheet, wb_source As Workbook)
    Dim add_columns_str As String
    Dim cols_arr As Variant
    Dim col As Variant
    Dim delim As String: delim = "||"
    Dim col_last As Integer
    
    add_columns_str = GetConfigParameterValueB("Add Dummy Columns To Export", wb_source)
    cols_arr = Split(add_columns_str, delim)
    
    For Each col In cols_arr
        'Debug.Print (col)
        col_last = FindLastColumnInRange(ws_target.Range("$1:$1")) ' pass 1st row to find the last used column in it
        ws_target.Columns(col_last + 1).Cells(1, 1).value = col
    Next
    
End Sub

Private Function FindLastColumnInRange(rng As Range, _
                        Optional What As String = "*", _
                        Optional LookIn As XlFindLookIn = xlValues, _
                        Optional LookAt As XlLookAt = xlPart, _
                        Optional SearchOrder As XlSearchOrder = xlByRows, _
                        Optional SearchDirection As XlSearchDirection = xlPrevious, _
                        Optional MatchCase As Boolean = False, _
                        Optional MatchByte As Boolean = False, _
                        Optional SearchFormat As Boolean = False) As Integer

    FindLastColumnInRange = rng.Find(What, rng.Cells(1, 1), LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat).Column
End Function

'returns True if the file was deleted
Private Function DeleteFile(ByVal FileToDelete As String) As Boolean
    Dim fso As New FileSystemObject, aFile As File
    
    If (fso.FileExists(FileToDelete)) Then
        Set aFile = fso.GetFile(FileToDelete)
        aFile.Delete
    End If
    
    DeleteFile = Not (fso.FileExists(FileToDelete))
End Function

Private Function CreateNewFile(file_name As String, Optional new_sheet_name As String = "Sheet1") As Workbook
    'expected file_type values: "metadata", "manifest"
    Dim path As String, cfg_row As Integer, new_file_name As String
    'Dim abort As Boolean: abort = False
    Dim wb As Workbook
    
    Set wb = Workbooks.Add(xlWBATWorksheet) 'create a new excel file with a single sheet
    wb.Sheets(1).Name = new_sheet_name 'rename the sheet of the workbook
    On Error GoTo Err_SaveAs
    wb.SaveAs (file_name) 'save new excel file to a specified folder with the given name
    On Error GoTo 0
    Set CreateNewFile = wb

    Exit Function
Err_SaveAs:
    wb.Close False
    Set CreateNewFile = Nothing
End Function

Private Function GetFileNameToSave()
    'expected file_type values: "metadata", "manifest"
    
    'Dim timepoint As String, spec_prep As String
    Dim date_received As String
    'Dim cfg_row As Integer
    Dim post_fix As String
    Dim ship_date As String
    
'    timepoint = GetTimepoint() 'get timepoint value
'    spec_prep = GetSpecimenPrepAbbr() 'get specimen abbreviation value
    ship_date = GetDateReceived()
    post_fix = GetConfigParameterValueB("File Name Post-fix") 'get post-fix value for created files
    
    If Not IsDate(ship_date) Then
        'verify the returned ship_date value
        ship_date = Date
    End If
    
    ship_date = Format(ship_date, "yyyy_mm_dd")
    
    GetFileNameToSave = "CHARM_COVID_Detection_" & ship_date & "_prepared_internally" & Trim(post_fix) & ".xlsx"
End Function

Public Sub ShowVersionMsg()
    MsgBox "CHARM COVID Detection Validation Tool - version #" & Version _
    & vbCrLf & vbCrLf _
    & "Please send any comment and questions about this tool to 'stas.rirak@mssm.edu'", _
    vbInformation, "CHARM COVID Detection Validation Tool"
End Sub

Public Sub OpenHelpLink()
    Dim url As String
    
    url = GetConfigParameterValue("Help document")
    ThisWorkbook.FollowHyperlink (url)
    
End Sub

Public Function ValidateTimepointValue(timepoint As String)
    Dim timepoint_validation_failed As Boolean
    'validate time point
    If Len(timepoint) > 4 Then
        'log error
        'AddLogEntry "Provided Timepoint value '" & timepoint & "' is too long (longer then 4 characters).", LogMsgType.Error
        timepoint_validation_failed = True
    End If
    'validate time point
    If Len(timepoint) < 3 Then
        'log error
        'AddLogEntry "Provided Timepoint value '" & timepoint & "' is too short (less then 2 characters).", LogMsgType.Error
        timepoint_validation_failed = True
    End If
    'validate time point
    If UCase(Left(timepoint, 1)) <> "T" Then
        'log error
        'AddLogEntry "The Timepoint value '" & timepoint & "' has to start with letter 'T'.", LogMsgType.Error
        timepoint_validation_failed = True
    End If
    'log message, if timepoint is valid
'    If Not timepoint_validation_failed Then
'        'AddLogEntry "The Timepoint value '" & timepoint & "' was recognized as valid.", LogMsgType.info
'    End If
    ValidateTimepointValue = timepoint_validation_failed
End Function

Public Function getTimeLength(tStart As Date, tEnd As Date) As String
    Dim mSeconds As Long, mHours As Long, mMinutes As Long
    Dim strTime As String
    
    mSeconds = DateDiff("s", tStart, tEnd)
    mHours = mSeconds \ 3600
    mMinutes = (mSeconds - (mHours * 3600)) \ 60
    mSeconds = mSeconds - ((mHours * 3600) + (mMinutes * 60))
    
    If mHours > 0 Then strTime = mHours & " hours "
    If mMinutes > 0 Then strTime = strTime & mMinutes & " minutes "
    strTime = strTime & mSeconds & " seconds "
    
    getTimeLength = strTime
End Function

Public Function strToBool(value As String) As Boolean
    On Error GoTo errLab
    
    value = Trim(value)
    strToBool = CBool(value)
    Exit Function
    
errLab:
    strToBool = False
End Function

Public Function strToOr(ParamArray vars() As Variant) As Boolean
    Dim var As Variant
    Dim out As Boolean
    
    For Each var In vars
        'Debug.Print i
        out = strToBool(CStr(var))
        'Debug.Print out
        If out Then Exit For
    Next
    strToOr = out
End Function

Public Function OrOfArray(vars() As Variant) As Boolean
    Dim var As Variant
    Dim out As Boolean
    
    For Each var In vars
        'Debug.Print i
        out = strToBool(CStr(var))
        'Debug.Print out
        If out Then Exit For
    Next
    OrOfArray = out
End Function

Public Function DetectionFileLoadedRows() As Integer
    DetectionFileLoadedRows = Worksheets(DetectionFileWrkSh).Cells(Rows.Count, 1).End(xlUp).Row
End Function

Sub ResetFilters(wks As Worksheet)

    On Error GoTo ErrHandler 'commented to test, need to be uncommented
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    If wks.AutoFilterMode Then
        wks.AutoFilter.ShowAllData
    End If
    GoTo ExitMark
    
ErrHandler:
    MsgBox Err.Description, vbCritical
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'returns dictionary providing mapping between actual and expected COVID detection values
Private Function GetDetectionValuesDict() As Scripting.dictionary
    Set GetDetectionValuesDict = GetDictItems("Detection_Values")
End Function

Private Function GetDictItems(dict_name As String) As Scripting.dictionary
    Dim col_num As Integer
    Dim dict As New Scripting.dictionary

    col_num = FindColNumberOfDictCategory(dict_name)
    If col_num > 0 Then
        Set dict = GetDictItemsPerColNum(col_num)
    Else
        Set dict = Nothing
    End If
    
    Set GetDictItems = dict
End Function

'searches for a given parameter name on the config page and returns the row number it was found on
Private Function FindColNumberOfDictCategory(categ_name As String, Optional wb As Workbook = Nothing) As Integer
    Dim ws_cfg As Worksheet
    
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(DictionaryWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(DictionaryWrkSheet)
    End If
    
    If IsError(Application.Match(categ_name, ws_cfg.Rows(1), 0)) Then
        FindColNumberOfDictCategory = 0
    Else
        FindColNumberOfDictCategory = Application.Match(categ_name, ws_cfg.Rows(1), 0)
    End If
    
End Function

Private Function GetDictItemsPerColNum(col_num As Integer, Optional wb As Workbook = Nothing) As dictionary
    Dim ws_cfg As Worksheet
    Dim dict_range As Range
    Dim dict As New Scripting.dictionary
    Dim cell As Range
        
    If wb Is Nothing Then
        Set ws_cfg = Worksheets(DictionaryWrkSheet)
    Else
        Set ws_cfg = wb.Worksheets(DictionaryWrkSheet)
    End If
    
    'get range based on the number of column provided as a parameter
    Set dict_range = ws_cfg.Columns(col_num)
    
    'loop through cells of the range
    For Each cell In dict_range.Cells
        If cell.Row > 1 Then 'proceed only if this is not a first (header) row
            If cell.Row > ws_cfg.UsedRange.Rows.Count Then
                Exit For
            End If
            'Debug.Print cell.Address & " - " & cell.Offset(0, 1).Address
            If Len(Trim(cell.Value2)) > 0 Then 'check the dictionary key value is not blank
                'Debug.Print cell.Value2 & " - " & cell.Offset(0, 1).Value2
                dict.Add cell.Value2, cell.offset(0, 1).Value2
            End If
        End If
    Next
   
    Set GetDictItemsPerColNum = dict
End Function

'converts provided detection_val name to an expected name of it based on the harcoded mapping dictionary
Function GetDetectionValue(detection_val As String, Optional DetectionValuesDict As dictionary = Nothing) As String
    Dim dictionary As Scripting.dictionary
    Dim out_val As String
    
    'if DetectionValuesDict was not supplied, get a fresh copy of it
    If DetectionValuesDict Is Nothing Then
        Set dictionary = GetDetectionValuesDict()
    Else
        Set dictionary = DetectionValuesDict
    End If
        
    If dictionary.Exists(UCase(detection_val)) Then
        out_val = dictionary(UCase(detection_val))
    Else
        out_val = detection_val 'return same value, if no match was found
    End If
    GetDetectionValue = out_val
End Function

Public Function VerifyDetectionValues()
    Dim rng As Range
    Dim rows_num As Integer
    Dim cell As Range
    Dim detect_vals_dict As dictionary
    Dim detection_col As String
    
    On Error GoTo ErrHandler 'commented to test, need to be uncommented
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Set detect_vals_dict = GetDetectionValuesDict()
    
    detection_col = GetConfigParameterValueB("DetectionFile_DetectionResluts_Column")
    With Worksheets(DetectionFileWrkSh)
        rows_num = DetectionFileLoadedRows()
        Set rng = .Range(detection_col & "2:" & detection_col & CStr(rows_num))
        
        For Each cell In rng
            cell.value = GetDetectionValue(cell.Value2, detect_vals_dict)
        Next
        
    End With

GoTo ExitMark
    
ErrHandler:
    MsgBox Err.Description, vbCritical
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

Public Sub PopulateValidationColumn(validation_name As String, _
                            ParamArray vars() As Variant)
    Dim rng As Range, cell As Range
    Dim rows_num As Integer
    Dim valid_col As String
    Dim varValues() As Variant, var As Variant
    Dim aInitialised As Boolean
    
    
    valid_col = GetConfigParameterValueB(validation_name)
    'valid_col = "V" ' for testing only
    
    If Len(Trim(valid_col)) = 0 Then
        ' provided validation column was not identified, exit sub
        Exit Sub
    End If
    
    With Worksheets(DetectionFileWrkSh)
        rows_num = DetectionFileLoadedRows()
        
        'get reference to a range to be populated
        Set rng = .Range(valid_col & "2:" & valid_col & CStr(rows_num))
        
        'clear the range first
        rng.ClearContents
        
        For Each cell In rng
            'Debug.Print (cell.Address)
            
            Select Case validation_name
                Case "Existing Timepoints For Subject"
                    cell.value = ExistingTimepointsForSubject(.Range(CStr(vars(0)) & cell.Row).Value2)
                Case "New Timepoints For Subject in Manifest"
                    cell.value = NewTimepointsForSubjectInManifest(.Range(CStr(vars(0)) & cell.Row).Value2)
                Case "Duplicated Timepoints for Subject in Manifest"
                    cell.value = DuplicatedEntriesInManifest(.Range(CStr(vars(0)) & cell.Row).Value2)
                Case "Same TimePoint Exists"
                    cell.value = SameTimepointExists(.Range(CStr(vars(0)) & cell.Row).Value2, .Range(CStr(vars(1)) & cell.Row).Value2)
                

                Case "NewTimepoint Is Not Most Recent"
                    cell.value = IsNotTimepointLatest(.Range(CStr(vars(0)) & cell.Row).Value2, .Range(CStr(vars(1)) & cell.Row).Value2, 1)
                Case "Duplicates in manifest"
                    cell.value = IIf(Len(Trim(CStr(.Range(CStr(vars(0)) & cell.Row).Value2))) = 0, False, True)
                Case "No COVID Detection"
                    cell.value = IIf(Len(Trim(CStr(.Range(CStr(vars(0)) & cell.Row).Value2))) = 0, True, False)
                Case "Invalid Timepoint Format"
                    cell.value = ValidateTimepointValue(.Range(CStr(vars(0)) & cell.Row).Value2)
                Case "No Participant ID"
                    cell.value = IIf(Len(Trim(CStr(.Range(CStr(vars(0)) & cell.Row).Value2))) = 0, True, False)
                Case "Total Validation Failed"
                    For Each var In vars
                        'Redim array size
                        If Not aInitialised Then
                            ReDim Preserve varValues(0)
                            aInitialised = True
                        Else
                            ReDim Preserve varValues(ArrLength(varValues))
                        End If
                        varValues(ArrLength(varValues) - 1) = .Range(CStr(var) & cell.Row).Value2
                    Next
                    cell.value = OrOfArray(varValues)
                    
                    'reset array to 0 length
                    ReDim varValues(0)
                    aInitialised = False
            End Select
            
        Next
        
    End With
End Sub

Public Function PopulateValidationColumns() As Boolean
    On Error GoTo ErrHandler 'commented to test, need to be uncommented
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Worksheets(DetectionFileWrkSh).Protect UserInterFaceOnly:=True 'adds ability to update the Detection_File sheet programmatically, while it is protected for users
    
    'Debug.Print (Now())
    
    PopulateValidationColumn "Existing Timepoints For Subject", GetConfigParameterValueB("DetectionFile_SubjectID_Column") '"A"
    PopulateValidationColumn "New Timepoints For Subject in Manifest", GetConfigParameterValueB("DetectionFile_SubjectID_Column") '"A"
    PopulateValidationColumn "Duplicated Timepoints for Subject in Manifest", GetConfigParameterValueB("DetectionFile_SubjectID_Column") '"A"
    
    PopulateValidationColumn "Same TimePoint Exists", GetConfigParameterValueB("Existing Timepoints For Subject"), GetConfigParameterValueB("DetectionFile_Timepoint_Column") '"F", "B"
    PopulateValidationColumn "NewTimepoint Is Not Most Recent", GetConfigParameterValueB("Existing Timepoints For Subject"), GetConfigParameterValueB("DetectionFile_Timepoint_Column") '"F", "B"
    
    PopulateValidationColumn "Duplicates in manifest", GetConfigParameterValueB("Duplicated Timepoints for Subject in Manifest") '"H"
    PopulateValidationColumn "No COVID Detection", GetConfigParameterValueB("DetectionFile_DetectionResluts_Column") '"C"
    
    PopulateValidationColumn "Invalid Timepoint Format", GetConfigParameterValueB("DetectionFile_Timepoint_Column") '"B"
    PopulateValidationColumn "No Participant ID", GetConfigParameterValueB("DetectionFile_SubjectID_Column") '"A"
    PopulateValidationColumn "Total Validation Failed", _
                                GetConfigParameterValueB("Same TimePoint Exists"), _
                                GetConfigParameterValueB("NewTimepoint Is Not Most Recent"), _
                                GetConfigParameterValueB("Duplicates in manifest"), _
                                GetConfigParameterValueB("No COVID Detection"), _
                                GetConfigParameterValueB("Invalid Timepoint Format"), _
                                GetConfigParameterValueB("No Participant ID")
                                '"I", "J", "K", "L", "M", "N"
    
    'Debug.Print (Now())
    
    PopulateValidationColumns = True
    
    GoTo ExitMark
    
ErrHandler:
    MsgBox "Unexpected error has occured during execution of 'PopulateValidationColumns' procedure: " & vbCrLf & Err.Description, vbCritical
    PopulateValidationColumns = False
ExitMark:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function


