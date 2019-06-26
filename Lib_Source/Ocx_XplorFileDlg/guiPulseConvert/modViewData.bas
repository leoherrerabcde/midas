Attribute VB_Name = "modViewData"
'---------------------------------------------------------------------------------------
' Module    : modViewData
' Author    : Leo Herrera
' Date      : 06/03/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_Index_Pls_Ini            As Long
Private PV_Index_Pls_End            As Long
Private PV_Count_Pls                As Long
Private PV_PlsList_Count            As Long
Private PV_IndexSheet               As Long
Private PV_IndexSpread              As Long
'Private PV_Header_Count             As Long
'Private PV_Hdr_Long_Count           As Long
'Private PV_Hdr_Double_Count         As Long



Function Init_Index(LstVw As ListView, _
                    IndexSpread As Long, _
                    IndexSheet As Long, _
                    IndexPulse As Long) As Boolean

    PV_Index_Pls_Ini = IndexPulse
    PV_Count_Pls = ListViewGetVisibleCount(LstVw)
    PV_Index_Pls_End = PV_Index_Pls_Ini + PV_Count_Pls - 1
    PV_IndexSpread = IndexSpread
    PV_IndexSheet = IndexSheet
    PV_PlsList_Count = Pulse_GetSheetPulseCount(IndexSpread, IndexSheet)
    If PV_Count_Pls < PV_PlsList_Count Then
        Init_Index = True
    Else
        Init_Index = False
    End If

End Function

'Sub Load_List_Data(LstVw As ListView)
'
'Dim lvAddItems          As Long
'Dim lArray()            As Long
'Dim dArray()            As Double
'Dim lvCount             As Long
'Dim lvIndexIni          As Long
'Dim lvIndexEnd          As Long
'Dim lvLstItem           As ListItem
'Dim i, j                As Long
'Dim lvHeaderCount       As Integer
'
'    With LstVw.ListItems
'        ValidateRect LstVw.hwnd, 0&
'
'        lvIndexIni = PV_Index_Ini
'        lvIndexEnd = lvIndexIni + PV_Count
'        lvCount = PV_Count
'
'        If PV_List_Count = 0 Then
'            Exit Sub
'        End If
'        If lvIndexEnd >= PV_List_Count Then
'            lvIndexEnd = PV_List_Count - 1
'            lvCount = lvIndexEnd - lvIndexIni
'        End If
'
'        lvAddItems = lvCount - .Count
'        If lvAddItems > 0 Then
'            lvHeaderCount = PV_Header_Count
'            For i = 1 To lvAddItems
'                Set lvLstItem = .Add
'                LstVw_AddLstSubItems lvLstItem.ListSubItems, lvHeaderCount
'            Next
'        Else
'            If lvAddItems < 0 Then
'                For i = 1 To -lvAddItems
'                    .Remove (.Count)
'                Next
'            End If
'        End If
'
'        ReDim lArray(PV_Hdr_Long_Count - 1)
'        ReDim dArray(PV_Hdr_Double_Count - 1)
'
'        For i = 1 To lvCount
'            Pulse_GetErrorPointer lvIndexIni + i - 1, lArray(0), dArray(0)
'            .Item(i).Text = GetErrorCode(lArray(0))
'            For j = 1 To PV_Hdr_Long_Count - 1
'                .Item(i).ListSubItems(j) = lArray(j)
'            Next
'            For j = 0 To PV_Hdr_Double_Count - 1
'                .Item(i).ListSubItems(PV_Hdr_Long_Count + j) = dArray(j)
'            Next
'        Next
'        InvalidateRect LstVw.hwnd, 0&, 0&
'        AutoAjusteColumnWidth LstVw
'    End With
'
'End Sub

Sub Load_Pwd_Data(LstVw As ListView)

Dim lv_Pls_Count            As Long
Dim lv_Index_Ini            As Long

    lv_Index_Ini = PV_Index_Pls_Ini
'    lv_Index_Ini = PV_Index_Pls_Ini - PV_Count_Pls / 2
'    If lv_Index_Ini < 0 Then
'        lv_Index_Ini = 0
'    End If
    If lv_Index_Ini + PV_Count_Pls <= PV_PlsList_Count - 1 Then
        lv_Pls_Count = PV_Count_Pls
    Else
        lv_Pls_Count = PV_PlsList_Count - lv_Index_Ini
    End If
    ShowSheetWindow LstVw, GV_ActualColumnConfig, _
                    PV_IndexSpread, _
                    PV_IndexSheet, _
                    lv_Index_Ini, _
                    lv_Pls_Count, _
                    PV_Index_Pls_Ini
    
End Sub

Sub ProcessScrollData(LstVw As ListView, _
                    IndexCmd As Integer)

Dim Count           As Long
    
    Count = ListViewGetVisibleCount(LstVw) '- 1
    Select Case IndexCmd
        Case Is = 1
            If PV_Index_Pls_Ini Then
                PV_Index_Pls_Ini = 0
                RefreshLstVwData LstVw
            End If
        Case Is = 2
            If PV_Index_Pls_Ini < Count Then
                If PV_Index_Pls_Ini <> 0 Then
                    PV_Index_Pls_Ini = 0
                    RefreshLstVwData LstVw
                End If
            Else
                PV_Index_Pls_Ini = PV_Index_Pls_Ini - Count
                RefreshLstVwData LstVw
            End If
        Case Is = 3
            If PV_Index_Pls_Ini + Count < PV_PlsList_Count - 1 Then
                PV_Index_Pls_Ini = PV_Index_Pls_Ini + Count
                RefreshLstVwData LstVw
            End If
        Case Is = 4
            If PV_Index_Pls_Ini <> PV_PlsList_Count - Count - 1 Then
                PV_Index_Pls_Ini = PV_PlsList_Count - Count - 1
                RefreshLstVwData LstVw
            End If
    End Select
    
End Sub
                    
Sub RefreshLstVwData(LstVw As ListView)

    Load_Pwd_Data LstVw
    
End Sub



