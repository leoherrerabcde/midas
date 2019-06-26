Attribute VB_Name = "modSheetIssue"
'---------------------------------------------------------------------------------------
' Module    : modSheetIssue
' Author    : Leo Herrera
' Date      : 18/01/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Sub SetActualColumnConfig(LVColCfg As typeConfigSheetColumns)
    
Dim i           As Integer

    With GV_ActualColumnConfig
        .ColumnConfigName = LVColCfg.ColumnConfigName
        If .Count <> LVColCfg.Count Then
            .Count = LVColCfg.Count
            ReDim .Column(.Count - 1)
        End If
        For i = 0 To .Count - 1
            .Column(i).ColumnName = LVColCfg.Column(i).ColumnName
            .Column(i).Order = LVColCfg.Column(i).Order
            .Column(i).Visible = LVColCfg.Column(i).Visible
        Next
    End With
    
End Sub

Sub ShowSheetWindow(LstVw As ListView, _
                    LV_ActualColumnConfig As typeConfigSheetColumns, _
                    ByVal IndexSpread As Long, _
                    ByVal IndexSheet As Long, _
                    ByVal IndexPulseIni As Long, _
                    ByVal PulseQty As Long, _
                    ByVal IndexHighLight As Long)

Dim lPulseCount     As Long
Dim i, j            As Long
Dim ldArray()       As Double
Dim InvalidateCount As Long
Dim Limit           As Long
Dim lCols           As Long
Dim lvInstance      As Long
Dim lvIniPreVw      As Long
Dim lvEndPreVw      As Long

    With LstVw
        lvIniPreVw = IndexPulseIni
        lvEndPreVw = lvIniPreVw + PulseQty - 1
        lCols = Pulse_Field_Count
        ReDim ldArray(lCols - 1)
        
        ValidateRect LstVw.hwnd, 0&
        ListViewSetListItems LstVw, PulseQty, lCols
        
        lPulseCount = Pulse_GetSheetPulseCount(IndexSpread, IndexSheet)
        If lvEndPreVw > lPulseCount Then
            lvEndPreVw = lPulseCount
        End If
        
        For i = lvIniPreVw To lvEndPreVw
            Pulse_GetPwd IndexSpread, IndexSheet, i, ldArray(0)
            LstVwIssueDoubleItemWithFilter LstVw, _
                                            i - lvIniPreVw + 1, _
                                            ldArray, _
                                            LV_ActualColumnConfig
                                            
            If i = IndexHighLight Then
                Set LstVw.SelectedItem = LstVw.ListItems(i - lvIniPreVw + 1)
                'LstVw.ListItems(i - lvIniPreVw + 1).Selected = True
                'LstVw.ListItems(i - lvIniPreVw + 1).Selected
            End If
            ValidateRect LstVw.hwnd, 0&
        Next
    End With
    InvalidateRect LstVw.hwnd, 0&, 0&
    AutoAjusteColumnWidth LstVw
    
End Sub



