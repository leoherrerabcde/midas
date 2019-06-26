Attribute VB_Name = "ModListViewFunctions"
' LHE 29 Sept 2006
' Archivo .bas Creado por LHE
' Rutinas de Manejo de ListView

Option Explicit



Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long, ByVal bErase As Long)
Declare Sub ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long)


Private Type POINT
   X As Long
   Y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETITEM As Long = LVM_FIRST + 5
Private Const LVM_FINDITEM As Long = LVM_FIRST + 13
Private Const LVM_ENSUREVISIBLE = LVM_FIRST + 19
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30
Private Const LVM_GETTOPINDEX = LVM_FIRST + 39
Private Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Private Const LVM_SETITEMSTATE As Long = LVM_FIRST + 43
Private Const LVM_GETITEMSTATE As Long = LVM_FIRST + 44
Private Const LVM_GETITEMTEXT As Long = LVM_FIRST + 45
Private Const LVM_SORTITEMS As Long = LVM_FIRST + 48
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = LVM_FIRST + 55
Private Const LVM_SETCOLUMNORDERARRAY = LVM_FIRST + 58
Private Const LVM_GETCOLUMNORDERARRAY = LVM_FIRST + 59

Private Const LVS_EX_GRIDLINES As Long = &H1
Private Const LVS_EX_SUBITEMIMAGES As Long = &H2
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_TRACKSELECT As Long = &H8
Private Const LVS_EX_HEADERDRAGDROP As Long = &H10
Private Const LVS_EX_FULLROWSELECT As Long = &H20

Private Const LVFI_PARAM As Long = 1

Private Const LVIF_TEXT As Long = 1
Private Const LVIF_IMAGE As Long = 2
Private Const LVIF_PARAM As Long = 4
Private Const LVIF_STATE As Long = 8
Private Const LVIF_INDENT As Long = &H10
Private Const LVIF_NORECOMPUTE As Long = &H800
Private Const LVIS_STATEIMAGEMASK As Long = &HF000&

Private Type LV_Item
   mask As Long
   Index As Long
   SubItem As Long
   state As Long
   StateMask As Long
   Text As String
   TextMax As Long
   Icon As Long
   Param As Long
   Indent As Long
End Type

Private Type LV_FINDINFO
   Flags As Long
   pSz As String
   lParam As Long
   pt As POINT
   vkDirection As Long
End Type
'--- ListView Set Column Width Messages ---'
Public Enum LVSCW_Styles
   LVSCW_AUTOSIZE = -1
   LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

Public Enum LVStylesEx
   Checkboxes = LVS_EX_CHECKBOXES
   FullRowSelect = LVS_EX_FULLROWSELECT
   GridLines = LVS_EX_GRIDLINES
   HeaderDragDrop = LVS_EX_HEADERDRAGDROP
   SubItemImages = LVS_EX_SUBITEMIMAGES
   TrackSelect = LVS_EX_TRACKSELECT
End Enum

'--- Sorting Variables ---'
Public Enum LVItemTypes
   lvDate = 0
   lvNumber = 1
   lvBinary = 2
   lvAlphabetic = 3
End Enum
Public Enum LVSortTypes
   lvAscending = 0
   lvDescending = 1
End Enum

Enum ImageSizingTypes
   [sizeNone] = 0
   [sizeCheckBox]
   [sizeIcon]
End Enum

'--- Array used to speed custom sorts ---'
Private m_lvSortData() As LV_Item
Private m_lvSortColl As Collection
Private m_lvSortColumn As Long
Private m_lvHWnd As Long
Private m_lvSortType As LVItemTypes
'
'
'


Public Function LVSetStyleEx(LV As ListView, ByVal NewStyle As LVStylesEx, ByVal NewVal As Boolean) As Boolean
   
   Dim nStyle As Long
   
   ' get the current ListView style
   nStyle = SendMessage(LV.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&)
   
   If NewVal Then
      ' set the extended style bit
      nStyle = nStyle Or NewStyle
   Else
      ' remove the extended style bit
      nStyle = nStyle Xor NewStyle
   End If
   
   ' set the new ListView style
   LVSetStyleEx = CBool(SendMessage(LV.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal nStyle))

End Function

Private Function LVCompareNumbers(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
   Static dat1 As Double
   Static dat2 As Double
   
   ' lookup text in listview based on index, and convert to double
   On Error Resume Next
   dat1 = CDbl(LVGetItemText(lParam1, m_lvHWnd))
   dat2 = CDbl(LVGetItemText(lParam2, m_lvHWnd))
   'On Error GoTo 0
   
   '--- this sorts ascending
   LVCompareNumbers = Sgn(dat1 - dat2)
   
   '--- this sorts descending
   If SortOrder = lvDescending Then
      LVCompareNumbers = -LVCompareNumbers
   End If
End Function

Private Function LVCompareText(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
   
   'Static dat1 As Date
   'Static dat2 As Date
   Dim LV_Str1  As String
   Dim LV_Str2  As String
   
   ' lookup text in listview based on index, and convert to date
   On Error Resume Next
   LV_Str1 = LVGetItemText(lParam1, m_lvHWnd)
   LV_Str2 = LVGetItemText(lParam2, m_lvHWnd)
   'On Error GoTo 0

   '--- this sorts ascending
   If LV_Str1 > LV_Str2 Then
        LVCompareText = 1
    ElseIf LV_Str1 < LV_Str2 Then
        LVCompareText = -1
    Else
        LVCompareText = 0
    End If
   
   '--- this sorts descending
   If SortOrder = lvDescending Then
      LVCompareText = -LVCompareText
   End If
   
End Function

Public Function LVGetItemText(lParam As Long, hwnd As Long) As String
   
   Dim objFind As LV_FINDINFO
   Dim Index As Long
   Dim objItem As LV_Item
   Dim nRet As Long
   
   ' Convert the input parameter to an index in the list view
   With objFind
      .Flags = LVFI_PARAM
      .lParam = lParam
   End With
   Index = SendMessage(hwnd, LVM_FINDITEM, -1, objFind)
   
   ' Obtain the name of the specified list view item
   With objItem
      .mask = LVIF_TEXT
      .SubItem = m_lvSortColumn
      .Text = Space(32)
      .TextMax = Len(.Text)
   End With
   
   ' Grab the text
   nRet = SendMessage(hwnd, LVM_GETITEMTEXT, Index, objItem)
   If nRet Then
      LVGetItemText = Left$(objItem.Text, nRet)
   End If
End Function

Function LstVw_Sort_by_Column(LV_LstVw As ListView, ByVal LV_ColumnHdr As MSComctlLib.ColumnHeader)

    With LV_LstVw
        .SortKey = LV_ColumnHdr.Index - 1
        .SortOrder = Abs(Not .SortOrder = 1)
        If .SortKey = 0 Then
            LVSortK LV_LstVw, .SortKey, lvNumber, .SortOrder
        ElseIf .SortKey = 6 Then
            LVSortK LV_LstVw, .SortKey, lvDate, .SortOrder
        Else
            LVSortK LV_LstVw, .SortKey, lvAlphabetic, .SortOrder
        End If
        'LVSortC LV_LstVw, .SortKey, lvNumber, .SortOrder
    End With

End Function

Public Function LVSortK(LV As ListView, ByVal Index As Long, ByVal ItemType As LVItemTypes, ByVal SortOrder As LVSortTypes) As Boolean
   'Dim tmr As New CStopWatch
   
   ' turn off the default sorting of the control
   With LV
      .Sorted = False
      .SortKey = Index
      .SortOrder = SortOrder
   End With

   ' store some values used during the sort
   m_lvSortColumn = Index
   m_lvSortType = ItemType
   m_lvHWnd = LV.hwnd
   'BuildLookup = 0
   
    If ItemType = lvNumber Then
        Call SendMessageLong(LV.hwnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompareNumbers)
    ElseIf ItemType = lvDate Then
        Call SendMessageLong(LV.hwnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompareDates)
    Else
        Call SendMessageLong(LV.hwnd, LVM_SORTITEMS, SortOrder, AddressOf LVCompareText)
    End If
    
End Function

Private Function LVCompareDates(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal SortOrder As Long) As Long
   Static dat1 As Date
   Static dat2 As Date
   
   ' lookup text in listview based on index, and convert to date
   On Error Resume Next
   dat1 = CDate(LVGetItemText(lParam1, m_lvHWnd))
   dat2 = CDate(LVGetItemText(lParam2, m_lvHWnd))
   'On Error GoTo 0

   '--- this sorts ascending
   LVCompareDates = Sgn(dat1 - dat2)
   
   '--- this sorts descending
   If SortOrder = lvDescending Then
      LVCompareDates = -LVCompareDates
   End If
End Function


Sub SaveColumnWideListView(ByRef lvLstView As ListView, _
                            ByRef lsApp As String, _
                            ByRef lsSection As String)

Dim ls_Key          As String
Dim i               As Integer

    With lvLstView
        For i = 1 To .ColumnHeaders.Count
            ls_Key = "Width Column " & Trim$(Str(i))
            SaveSetting lsApp, lsSection, ls_Key, .ColumnHeaders(i).Width
        Next
    End With
    

End Sub


Function LstVwFindItemChecked(LV_LstVw As ListView, Optional lvFindFirst As Boolean) As Integer

Dim i           As Integer
Dim LV_Checked  As Integer

    LV_Checked = 0
    With LV_LstVw
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked = True Then
                If LV_Checked Then
                    LstVwFindItemChecked = 0
                    Exit Function
                Else
                    LV_Checked = i
                    If lvFindFirst = True Then
                        Exit For
                    End If
                End If
            End If
            
        Next
        LstVwFindItemChecked = LV_Checked
    End With
    
End Function

Sub LoadColumnWideListView(ByRef lvLstView As ListView, _
                            ByRef lsApp As String, _
                            ByRef lsSection As String)

Dim ls_Key          As String
Dim i               As Integer
Dim lWidth          As Integer
'Dim lsWidth         As String

    On Error Resume Next
    With lvLstView
        For i = 1 To .ColumnHeaders.Count
            ls_Key = "Width Column " & Trim$(Str(i))
            lWidth = ChangeRegionalConfig(GetSetting(lsApp, lsSection, ls_Key, .ColumnHeaders(i).Width))
            .ColumnHeaders(i).Width = lWidth
        Next
    End With
    'On Error GoTo 0

End Sub

Function LstVw_Add_ListSubItems(ByRef lvLstVw As ListView, _
                    ByRef lsData() As String, Index As Long)

Dim i               As Integer

    With lvLstVw.ListItems(Index).ListSubItems
        For i = 0 To UBound(lsData)
            If .Count < i + 1 Then
                .Add , , lsData(i)
            Else
                .Item(i + 1).Text = lsData(i)
            End If
        Next
    End With

End Function

Sub LstVw_Add_ListItems(ByRef lvLstVw As ListView, _
                    ByRef lsData() As String, _
                    Optional ByVal VerPrimero As Boolean)

Dim i               As Integer
        
    ValidateRect lvLstVw.hwnd, 0&
    
    With lvLstVw
        For i = 0 To UBound(lsData)
            .ListItems.Add , , lsData(i)
        Next
        If VerPrimero = False Then
            .ListItems(.ListItems.Count).EnsureVisible
        Else
            .ListItems(1).EnsureVisible
        End If
    End With
        
End Sub

Sub AddDoubleItemListViewWithFilter(ByRef lvLstVw As ListView, _
                    ByRef lData() As Double, _
                    LV_ColumnConfig As typeConfigSheetColumns, _
                    Optional ByVal FirstColumnValue As Long, _
                    Optional ByVal VerPrimero As Boolean, _
                    Optional ByVal Index As Integer)
                    
Dim LV_LstSubItems  As ListSubItems
Dim i               As Integer
Dim i_max           As Integer

    'Set LV_LstSubItems = lvLstVw.ListItems.Add(, , FirstColumnValue)
    Set LV_LstSubItems = lvLstVw.ListItems.Add
    
    lvLstVw.ListItems(lvLstVw.ListItems.Count).Tag = Index
    
    ValidateRect lvLstVw.hwnd, 0&
    
    If VerPrimero = False Then
        lvLstVw.ListItems(lvLstVw.ListItems.Count).EnsureVisible
    End If
    
    
    With LV_LstSubItems
        i_max = 0
        LstVw_AddLstSubItems LV_LstSubItems, lvLstVw.ColumnHeaders.Count - 1
        For i = 0 To UBound(lData)
            If LV_ColumnConfig.Column(i).Visible = True Then
                If LV_ColumnConfig.Column(i).Order Then
                    LV_LstSubItems(LV_ColumnConfig.Column(i).Order).Text = _
                            lData(i)
                Else
                    lvLstVw.ListItems(lvLstVw.ListItems.Count).Text = lData(i)
                End If
                If i_max < LV_ColumnConfig.Column(i).Order + 1 Then
                    i_max = LV_ColumnConfig.Column(i).Order + 1
                End If
            End If
        Next
    End With

End Sub

Sub LstVwIssueDoubleItemWithFilter(ByRef lvLstVw As ListView, _
                    IndexLstItm As Long, _
                    ByRef lData() As Double, _
                    LV_ColumnConfig As typeConfigSheetColumns)
                    
Dim LV_LstSubItems  As ListSubItems
Dim i               As Integer
Dim i_max           As Integer

    Set LV_LstSubItems = lvLstVw.ListItems(IndexLstItm).ListSubItems
    
    'lvLstVw.ListItems(IndexLstItm).Tag = Index
    
    ValidateRect lvLstVw.hwnd, 0&
    
    With LV_LstSubItems
        i_max = 0
        'LstVw_AddLstSubItems LV_LstSubItems, lvLstVw.ColumnHeaders.Count - 1
        For i = 0 To UBound(lData)
            If LV_ColumnConfig.Column(i).Visible = True Then
                If LV_ColumnConfig.Column(i).Order Then
                    LV_LstSubItems(LV_ColumnConfig.Column(i).Order).Text = _
                            lData(i)
                Else
                    lvLstVw.ListItems(IndexLstItm).Text = lData(i)
                End If
                If i_max < LV_ColumnConfig.Column(i).Order + 1 Then
                    i_max = LV_ColumnConfig.Column(i).Order + 1
                End If
            End If
        Next
    End With

End Sub

Sub LstVw_AddLstSubItems(LV_LstSubItems As ListSubItems, Count As Integer)
    
Dim i           As Integer

    For i = 1 To Count
        LV_LstSubItems.Add
    Next
    
End Sub

Sub AddDoubleItemListView(ByRef lvLstVw As ListView, _
                    ByRef lData() As Double, _
                    Optional ByVal FirstColumnValue As Long, _
                    Optional ByVal VerPrimero As Boolean, _
                    Optional ByVal Index As Integer)
                    
Dim LV_LstSubItems  As ListSubItems
Dim i               As Integer
    
    'Set LV_LstSubItems = lvLstVw.ListItems.Add(, , FirstColumnValue)
    Set LV_LstSubItems = lvLstVw.ListItems.Add(, , FirstColumnValue)
    
    lvLstVw.ListItems(lvLstVw.ListItems.Count).Tag = Index
    
    ValidateRect lvLstVw.hwnd, 0&
    
    If VerPrimero = False Then
        lvLstVw.ListItems(lvLstVw.ListItems.Count).EnsureVisible
    End If
    
    With LV_LstSubItems
        For i = 0 To UBound(lData)
            If i Then
                .Add , , lData(i)
            Else
                lvLstVw.ListItems(lvLstVw.ListItems.Count).Text = lData(i)
            End If
        Next
    End With
    
End Sub

Sub AddItemListView(ByRef lvLstVw As ListView, _
                    ByRef lsData() As String, _
                    Optional ByVal VerPrimero As Boolean, _
                    Optional ByVal Index As Integer)
                    
Dim LV_LstSubItems  As ListSubItems
Dim i               As Integer
    
    Set LV_LstSubItems = lvLstVw.ListItems.Add(, , lsData(0))
    
    lvLstVw.ListItems(lvLstVw.ListItems.Count).Tag = Index
    
    ValidateRect lvLstVw.hwnd, 0&
    
    If VerPrimero = False Then
        lvLstVw.ListItems(lvLstVw.ListItems.Count).EnsureVisible
    End If
    
    With LV_LstSubItems
        For i = 1 To UBound(lsData)
            .Add , , lsData(i)
        Next
    End With
    
End Sub

Sub AddColumListView(ByRef lvLstVw As ListView, _
                    ByRef lsData() As String)
                    
Dim LV_LstSubItems   As ListSubItems
Dim i               As Integer
    
    With lvLstVw
        .ColumnHeaders.Clear
        For i = 0 To UBound(lsData)
            .ColumnHeaders.Add , , lsData(i)
        Next
        .ListItems.Clear
    End With
    
End Sub

Function LstVw_AddColumnHeader(LstVw As ListView, Count As Integer)

Dim i           As Integer
    
    For i = 1 To Count
        LstVw.ColumnHeaders.Add
    Next
    
End Function

Function LstVw_RemoveColumnsStartingFrom(Index As Integer, LstVw As ListView)

Dim i           As Integer

    With LstVw.ColumnHeaders
        For i = .Count To Index Step -1
            .Remove i
        Next
    End With
    
End Function

Function LstVw_Set_CheckBox(ByRef LstVw As ListView, ByVal lvValue As Boolean) As Long

Dim i           As Integer
'Dim LV_Count    As Long

    'LV_Count = 0
    With LstVw
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = lvValue
        Next
        If lvValue = True Then
            LstVw_Set_CheckBox = .ListItems.Count
        Else
            LstVw_Set_CheckBox = 0
        End If
    End With

End Function

Sub LstVw_Check_Selected(ByRef LstVw As ListView, Optional LvFnInv As Boolean)

    With LstVw
        If LvFnInv = True Then
            If .SelectedItem.Checked = True Then
                .SelectedItem.Checked = False
            Else
                .SelectedItem.Checked = True
            End If
        Else
            .SelectedItem.Checked = True
        End If
    End With
    
End Sub

Function LstVw_CopyColumnHeaders(ByRef LstVwDst As ListView, _
                                ByRef LstVwSrc As ListView, _
                                Optional lsHeader As String = "")

Dim i           As Integer

    LstVwDst.ListItems.Clear
    LstVwDst.ColumnHeaders.Clear
    
    If lsHeader <> "" Then
        LstVwDst.ColumnHeaders.Add , , lsHeader
    End If
    For i = 1 To LstVwSrc.ColumnHeaders.Count
        LstVwDst.ColumnHeaders.Add , , LstVwSrc.ColumnHeaders(i).Text
    Next
    
End Function

Function LstVw_Move_Chequed_To(ByRef LstVwSrc As ListView, ByRef LstVwDst As ListView) As Long

Dim i           As Integer
'Dim LV_Count    As Long

    'LV_Count = 0
    i = 1
    With LstVwSrc
        Do
            If .ListItems(i).Checked = True Then
                LstVwDst.ListItems.Add , , .ListItems(i).Text
                .ListItems.Remove i
            Else
                i = i + 1
            End If
        Loop Until i > .ListItems.Count
    End With

End Function

Function LstVw_Move_Selected_To(ByRef LstVwSrc As ListView, ByRef LstVwDst As ListView) As Long

Dim i           As Integer
'Dim LV_Count    As Long

    'LV_Count = 0
    i = 1
    With LstVwSrc
        Do
            If .ListItems(i).Selected = True Then
                LstVwDst.ListItems.Add , , .ListItems(i).Text
                .ListItems.Remove i
            Else
                i = i + 1
            End If
        Loop Until i > .ListItems.Count
    End With

End Function

Function Contar_Items_Chequeados(ByRef LstVw As ListView) As Long

Dim i           As Integer
Dim LV_Count    As Long

    LV_Count = 0
    With LstVw
        For i = 1 To .ListItems.Count
            If LVItemChecked(LstVw, i) = True Then
                LV_Count = LV_Count + 1
            End If
        Next
    End With
    
    Contar_Items_Chequeados = LV_Count
    
End Function

Function FormatHora(lsHora As String) As String

    FormatHora = Format(lsHora, "hh:mm:ss")
    
End Function

Function GetSubItems(ByRef LstVw As ListView, Index As Integer) As String

    GetSubItems = LstVw.SelectedItem.ListSubItems(Index).Text
    
End Function


    

Function ChangeRegionalConfig(lsNumber As String) As String

        If IsNumeric(lsNumber) = False Then
            ChangeRegionalConfig = Replace(lsNumber, ".", ",")
        Else
            ChangeRegionalConfig = lsNumber
        End If

End Function

Function GetIntegerFromStr(lsNumber As String) As Double

    If IsNumeric(GetIntegerFromStr) = True Then
        If lsNumber < 32768 Then
            GetIntegerFromStr = lsNumber
        Else
            GetIntegerFromStr = ChangeRegionalConfig(lsNumber)
        End If
    Else
        GetIntegerFromStr = ChangeRegionalConfig(lsNumber)
    End If
    
End Function

Public Sub LVSetColWidth(LV As ListView, ByVal ColumnIndex As Long, ByVal Style As LVSCW_Styles)
   '------------------------------------------------------------------------------
   '--- If you include the header in the sizing then the last column will
   '--- automatically size to fill the remaining listview width.
   '------------------------------------------------------------------------------
   With LV
      ' verify that the listview is in report view and that the column exists
      If .View = lvwReport Then
         If ColumnIndex >= 1 And ColumnIndex <= .ColumnHeaders.Count Then
            Call SendMessage(.hwnd, LVM_SETCOLUMNWIDTH, ColumnIndex - 1, ByVal Style)
         End If
      End If
   End With
End Sub

Public Function LVItemChecked(LV As ListView, ByVal Index As Long) As Boolean
   Dim nRet As Long
   Const MaskBit As Long = &H1000   '(2 ^ 12)
   
   ' get current statemask bits
   nRet = SendMessage(LV.hwnd, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_STATEIMAGEMASK)
   
   ' return what the Checked bit is set to
   LVItemChecked = (((nRet \ MaskBit) - 1) <> 0)
End Function

Public Sub LVSetAllColWidths(LV As ListView, ByVal Style As LVSCW_Styles)
   Dim ColumnIndex As Long
   '--- loop through all of the columns in the listview and size each
   With LV
      For ColumnIndex = 1 To .ColumnHeaders.Count
         LVSetColWidth LV, ColumnIndex, Style
      Next ColumnIndex
   End With
End Sub


Sub SetHighlightColumn(LV As ListView, _
                               clrHighlight As SystemColorConstants, _
                               clrDefault As OLE_COLOR, _
                               nColumn As Long, _
                               nSizingType As ImageSizingTypes, _
                               Picture1)

   Dim cnt     As Long  'counter
   Dim cl      As Long  'columnheader left
   Dim cw      As Long  'columnheader width
         
   On Local Error GoTo SetHighlightColumn_Error
   
   If LV.View = lvwReport Then
   
     'set up the listview properties
      With LV
        .Picture = Nothing  'clear picture
        .Refresh
        .Visible = 1
        .PictureAlignment = lvwTile
      End With  ' lv
        
     'set up the picture box properties
      With Picture1
         .AutoRedraw = False       'clear/reset picture
         .Picture = Nothing
         .BackColor = clrDefault
         .Height = 1
         .AutoRedraw = True        'assure image draws
         .BorderStyle = vbBSNone   'other attributes
         .ScaleMode = vbTwips
         '.Top = Form1.Top - 10000  'move it off screen
         .Visible = False
         .Height = 1               'only need a 1 pixel high picture
         .Width = Screen.Width
            
        'draw a box in the highlight colour
        'at location of the column passed
         cl = LV.ColumnHeaders(nColumn).Left
         cw = LV.ColumnHeaders(nColumn).Left + _
              LV.ColumnHeaders(nColumn).Width
         Picture1.Line (cl, 0)-(cw, 210), clrHighlight, BF
         
         .AutoSize = True
      End With  'Picture1
     
     'set the lv picture to the
     'Picture1 image
      LV.Refresh
      LV.Picture = Picture1.Image
      
   Else
    
      LV.Picture = Nothing
        
   End If  'lv.View = lvwReport

SetHighlightColumn_Exit:
On Local Error GoTo 0
Exit Sub
    
SetHighlightColumn_Error:

  'clear the listview's picture and exit
   With LV
      .Picture = Nothing
      .Refresh
   End With
   
   Resume SetHighlightColumn_Exit
    
End Sub


Sub AutoAjusteColumnWidth(ByRef LstVw As ListView)

   LstVw.Visible = False
   Call LVSetAllColWidths(LstVw, LVSCW_AUTOSIZE_USEHEADER)
   LstVw.Visible = True

End Sub

Public Function ListViewGetVisibleCount(LV As ListView) As Long
   ListViewGetVisibleCount = SendMessage(LV.hwnd, LVM_GETCOUNTPERPAGE, 0&, _
       ByVal 0&)
End Function

Public Sub ListView_AddLstItems(LstVw As ListView, Count As Long, Cols As Long)

Dim i, j                As Long
Dim LV_LstSbItm         As ListSubItems

    For i = 1 To Count
        Set LV_LstSbItm = LstVw.ListItems.Add
        LstVw_AddLstSubItems LV_LstSbItm, Cols - 1
    Next
    
End Sub

Public Sub ListView_RemoveLstItems(LstVw As ListView, Count As Long)

Dim i           As Long

    For i = 1 To Count
        If LstVw.ListItems.Count Then
            LstVw.ListItems.Remove LstVw.ListItems.Count
        End If
    Next
    
End Sub

Public Sub ListViewSetListItems(LstVw As ListView, Count As Long, Cols As Long)

Dim lv_ItemCount        As Long

    With LstVw.ListItems
        lv_ItemCount = Count - .Count
        If lv_ItemCount > 0 Then
            ListView_AddLstItems LstVw, lv_ItemCount, Cols
        Else
            If lv_ItemCount < 0 Then
                ListView_RemoveLstItems LstVw, -lv_ItemCount
            End If
        End If
    End With
    
End Sub
