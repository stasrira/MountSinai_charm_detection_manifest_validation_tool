Attribute VB_Name = "mdlMain"
Option Explicit

Public Const Version = "1.00"

Private Const ExistingDetectionsWrkSh = "COVID_Detection_Existing"
Private Const DetectionFileWrkSh = "Detection_File"
Private Const ConfigWrkSheet = "config"

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
    
    subject_id_col = GetConfigParameterValueB("DetectionFile_SubjectID_Column")
    timepoint_offset_num = GetConfigParameterValueB("DetectionFile_TimepointColumn_Offset_From_Subject")
    
    Set rng = Worksheets(DetectionFileWrkSh).Range(subject_id_col & ":" & subject_id_col) 'COVID_Detection_Existing '"A:A"
    
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
    
    subject_id_col = GetConfigParameterValueB("DetectionFile_SubjectID_Column")
    timepoint_offset_num = GetConfigParameterValueB("DetectionFile_TimepointColumn_Offset_From_Subject")
    
    Set rng = Worksheets(DetectionFileWrkSh).Range(subject_id_col & ":" & subject_id_col) 'COVID_Detection_Existing '"A:A"
    
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

Function checkConditionalFormat(sheet As String, column_letter As String, condition_val As String) As Boolean
    Dim ws As Worksheet
    Dim col As Range
    
    Set ws = Worksheets(sheet)
    Set col = ws.Range(column_letter & ":" & column_letter)
    
    checkConditionalFormat = IsNumeric(Application.Match(condition_val, col, 0))
    
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
