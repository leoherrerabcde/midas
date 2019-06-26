VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmErrorView 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9885
   Begin VB.Frame FrameViewErrors 
      Caption         =   "Visualización de Errores"
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.PictureBox pictureRight 
         BorderStyle     =   0  'None
         Height          =   460
         Left            =   5040
         ScaleHeight     =   465
         ScaleWidth      =   4215
         TabIndex        =   8
         Top             =   120
         Width           =   4215
         Begin MSComctlLib.Toolbar toolbarControles 
            Height          =   390
            Index           =   1
            Left            =   1800
            TabIndex        =   10
            Top             =   40
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            ImageList       =   "ImageListControles"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   4
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label lblPreView 
            Caption         =   "Pre Visualizacion:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   1695
         End
      End
      Begin MSComctlLib.ImageList ImageListControles 
         Left            =   4440
         Top             =   3000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmErrorView.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmErrorView.frx":0A12
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmErrorView.frx":1424
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmErrorView.frx":19BE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar toolbarControles 
         Height          =   390
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageListControles"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Timer tmrPointer 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   4560
         Top             =   4200
      End
      Begin VB.PictureBox picSplitter 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         FillColor       =   &H00808080&
         Height          =   2100
         Left            =   4830
         ScaleHeight     =   914.43
         ScaleMode       =   0  'User
         ScaleWidth      =   624
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.CommandButton cmdUpDown 
         Height          =   255
         Index           =   1
         Left            =   3840
         Picture         =   "frmErrorView.frx":1F58
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUpDown 
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView LstVwPWD 
         Height          =   5295
         Left            =   5040
         TabIndex        =   2
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LstVwErrorList 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Image imgSplitter 
         Height          =   2100
         Left            =   4560
         MousePointer    =   9  'Size W E
         Top             =   870
         Width           =   45
      End
      Begin VB.Label lbl 
         Caption         =   "Lista de Errores:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmErrorView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmErrorView
' Author    : Leo Herrera
' Date      : 13/01/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_Index_Err_Ini            As Long
Private PV_Index_Err_End            As Long
Private PV_Count_Err                As Long
Private PV_ErrList_Count            As Long

Private PV_Index_Pls_Ini            As Long
Private PV_Index_Pls_End            As Long
Private PV_Count_Pls                As Long
Private PV_PlsList_Count            As Long
Private PV_IndexSheet               As Long
Private PV_IndexSpread              As Long

Private PV_Header_Count             As Long
Private PV_Hdr_Long_Count           As Long
Private PV_Hdr_Double_Count         As Long

Private objIzq As ListView
Private objDer As ListView
'
Private moviendo As Boolean
Private Const splitLimit As Long = 15&
'
Private clsMoveSplit        As clsWndSplit

Sub SetPwdHeaders(LstVwOutput As ListView)

Dim i               As Long
Dim Index           As Long
Dim IndexMaxOut     As Long
Dim IndexMaxHide    As Long
Dim LV_ColumnHeader As ColumnHeaders

    With GV_ActualColumnConfig
        For i = 0 To .Count - 1
            Index = .Column(i).Order + 1
            If .Column(i).Visible = True Then
                Set LV_ColumnHeader = LstVwOutput.ColumnHeaders
                If IndexMaxOut < Index Then
                    IndexMaxOut = Index
                End If
                LV_ColumnHeader.Item(Index).Text = .Column(i).ColumnName
                LV_ColumnHeader.Item(Index).Tag = i
            End If
        Next
    End With

End Sub


'Private Sub sizeControls(ByVal X As Long, FrameContainer As Frame)
'
'Dim tMinWidth As Long
'    '
'    On Error Resume Next
'    '
'    ' el ancho mínimo que tendrá cada panel
'    tMinWidth = Screen.TwipsPerPixelY * 90
'    '
'    ' asignar el ancho
'    If X < tMinWidth Then X = tMinWidth
'    If X > (FrameContainer.Width - tMinWidth) Then X = FrameContainer.Width - tMinWidth
'    objIzq.Width = X - objIzq.Left - imgSplitter.Width
'    imgSplitter.Left = X
'    SaveSetting App.Title, _
'                Me.Name, _
'                Me.imgSplitter.Name & ".Left", _
'                Me.imgSplitter.Left
'    objDer.Left = X + imgSplitter.Width
'    objDer.Width = FrameContainer.Width - objDer.Left - objIzq.Left '140)
'
'    Me.lblPreView.Left = X
'    Me.toolbarControles(1).Left = Me.lblPreView.Left + Me.lblPreView.Width
'
'    imgSplitter.Top = objIzq.Top
'    imgSplitter.Height = objIzq.Height
'
'End Sub

Sub Load_Error_List(LstVw As ListView)

Dim lvAddItems          As Long
Dim lArray()            As Long
Dim dArray()            As Double
Dim lvCount             As Long
Dim lvIndexIni          As Long
Dim lvIndexEnd          As Long
Dim lvLstItem           As ListItem
Dim i, j                As Long
Dim lvHeaderCount       As Integer

    With LstVw.ListItems
        ValidateRect LstVw.hwnd, 0&
        
        lvIndexIni = PV_Index_Err_Ini
        lvIndexEnd = lvIndexIni + PV_Count_Err
        lvCount = PV_Count_Err
        
        If PV_ErrList_Count = 0 Then
            Exit Sub
        End If
        If lvIndexEnd >= PV_ErrList_Count Then
            lvIndexEnd = PV_ErrList_Count - 1
            lvCount = lvIndexEnd - lvIndexIni
        End If
            
        lvAddItems = lvCount - .Count
        If lvAddItems > 0 Then
            lvHeaderCount = PV_Header_Count
            For i = 1 To lvAddItems
                Set lvLstItem = .Add
                LstVw_AddLstSubItems lvLstItem.ListSubItems, lvHeaderCount
            Next
        Else
            If lvAddItems < 0 Then
                For i = 1 To -lvAddItems
                    .Remove (.Count)
                Next
            End If
        End If
        
        ReDim lArray(PV_Hdr_Long_Count - 1)
        ReDim dArray(PV_Hdr_Double_Count - 1)
        
        For i = 1 To lvCount
            Pulse_GetErrorPointer lvIndexIni + i - 1, lArray(0), dArray(0)
            .Item(i).Text = GetErrorCode(lArray(0))
            For j = 1 To PV_Hdr_Long_Count - 1
                .Item(i).ListSubItems(j) = lArray(j)
            Next
            For j = 0 To PV_Hdr_Double_Count - 1
                .Item(i).ListSubItems(PV_Hdr_Long_Count + j) = dArray(j)
            Next
        Next
        InvalidateRect LstVw.hwnd, 0&, 0&
        AutoAjusteColumnWidth LstVw
    End With
    
End Sub

Sub Init_Error_Index(LstVw As ListView)

    PV_Index_Err_Ini = 0
    PV_Count_Err = ListViewGetVisibleCount(LstVw) - 1
    PV_Index_Err_End = PV_Index_Err_Ini + PV_Count_Err - 1
    PV_Count_Err = PV_Index_Err_End - PV_Index_Err_Ini + 1
    
    PV_ErrList_Count = Pulse_GetErrorListCount
    Pulse_GetErrorFieldCount PV_Hdr_Long_Count, PV_Hdr_Double_Count
    PV_Header_Count = PV_Hdr_Long_Count + PV_Hdr_Double_Count
    
End Sub

Sub Init_Pulse_Index(LstVw As ListView, _
                    IndexSpread As Long, _
                    IndexSheet As Long, _
                    IndexPulse As Long)

    PV_Index_Pls_Ini = IndexPulse
    PV_Count_Pls = ListViewGetVisibleCount(LstVw)
    PV_Index_Pls_End = PV_Index_Pls_Ini + PV_Count_Pls - 1
    PV_IndexSpread = IndexSpread
    PV_IndexSheet = IndexSheet
    PV_PlsList_Count = Pulse_GetSheetPulseCount(IndexSpread, IndexSheet)
    If PV_Count_Pls < PV_PlsList_Count Then
        Me.toolbarControles(1).Enabled = True
        Me.toolbarControles(1).Visible = True
    Else
        Me.toolbarControles(1).Enabled = False
        Me.toolbarControles(1).Visible = False
    End If

End Sub

Sub Load_Error_Header(LstVw As ListView)

Dim lsHeader            As String
Dim i                   As Long

    With LstVw.ColumnHeaders
        .Clear
        lsHeader = Space(250)
        For i = 0 To PV_Header_Count - 1
            Pulse_GetErrorFieldHeader i, lsHeader
            .Add , , lsHeader
        Next
        AutoAjusteColumnWidth LstVw
    End With
    
End Sub

Sub Set_Form_Caption()

'    If m_Project.IsNewProject = True Then
'        Me.Caption = "Proyecto Nuevo - Ubicación de Pulsos"
'    Else
'        Me.Caption = "Seleccionando nueva ubicación de Pulsos"
'    End If
    
End Sub

Private Sub cmdUpDown_Click(Index As Integer)

    With Me
        Select Case Index
            Case Is = 0
                If PV_Index_Err_Ini < PV_Count_Err Then
                    If PV_Index_Err_Ini <> 0 Then
                        PV_Index_Err_Ini = 0
                        .Load_Error_List .LstVwErrorList
                    End If
                Else
                    PV_Index_Err_Ini = PV_Index_Err_Ini - PV_Count_Err
                    .Load_Error_List .LstVwErrorList
                End If
            Case Is = 1
                If PV_Index_Err_Ini + PV_Count_Err < PV_ErrList_Count Then
                    PV_Index_Err_Ini = PV_Index_Err_Ini + PV_Count_Err
                    .Load_Error_List .LstVwErrorList
                End If
        End Select
    End With
    
End Sub

Private Sub Form_Load()

    With Me
        Set clsMoveSplit = New clsWndSplit
        clsMoveSplit.Constructor .imgSplitter, .picSplitter, .FrameViewErrors, _
        .LstVwErrorList, .LstVwPWD, Me, .pictureRight, 120
        '.toolbarControles(0).Left + toolbarControles(0).Width
        
        GV_Mdi.Set_Status_MnuProject VerifyErrors, False
        m_Project.LoadWorkSpace
        Set objIzq = .LstVwErrorList
        Set objDer = .LstVwPWD
        Set_Form_Caption
        
        .imgSplitter.Left = GetSetting(App.Title, Me.Name, .imgSplitter.Name & ".Left", .imgSplitter.Left)
        SaveSetting App.Title, Me.Name, .imgSplitter.Name & ".Left", .imgSplitter.Left
        
        .LstVwErrorList.ListItems.Clear
        .LstVwErrorList.ColumnHeaders.Clear
        Init_Error_Index .LstVwErrorList
        Load_Error_Header .LstVwErrorList
        Load_Error_List .LstVwErrorList
        LstVw_AddColumnHeader .LstVwPWD, GV_ActualColumnConfig.Count
        SetPwdHeaders .LstVwPWD
        ShowPulseList
    End With
    
    Set_MousePointer vbDefault

End Sub

Private Sub Form_Resize()

Dim lvRelPos            As Double

    With Me
        'lvrelpos=()/.FrameViewErrors.Width
        .FrameViewErrors.Width = .ScaleWidth - .FrameViewErrors.Left
        '.LstVwPWD.Width = .FrameViewErrors.Width - .LstVwPWD.Left - .LstVwErrorList.Left
    
        clsMoveSplit.sizeControls imgSplitter.Left
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    m_Project.DiscardWorkSpace
    GV_clsErrorView.ClearLoaded
    
End Sub

'Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    With imgSplitter
'        picSplitter.Move .Left, .Top, .Width \ 3, .Height - 20
'    End With
'    picSplitter.Visible = True
'    moviendo = True
'
'End Sub
'
'Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'Dim sglPos As Single
'    '
'
'    If moviendo Then
'        sglPos = X + imgSplitter.Left
'        If sglPos < splitLimit Then
'            picSplitter.Left = splitLimit
'        ElseIf sglPos > Me.Width - splitLimit Then
'            picSplitter.Left = Me.Width - splitLimit
'        Else
'            picSplitter.Left = sglPos
'        End If
'    End If
'
'End Sub
'
'Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    With Me
'        sizeControls picSplitter.Left, .FrameViewErrors
'        picSplitter.Visible = False
'    End With
'    moviendo = False
'
'End Sub

Private Sub LstVwErrorList_Click()

    Me.ShowPulseList
    
End Sub

Sub ShowPulseList()

Dim Index           As Long
Dim IndexPulse      As Long
Dim IndexSheet      As Long
Dim IndexSpread     As Long

    With Me
        If .LstVwErrorList.ListItems.Count = 0 Then
            Exit Sub
        End If
        .toolbarControles(1).Enabled = True
        With .LstVwErrorList.SelectedItem.ListSubItems
            IndexSpread = .Item(1).Text
            IndexSheet = .Item(2).Text
            IndexPulse = .Item(3).Text
        End With
        .Init_Pulse_Index .LstVwPWD, IndexSpread, IndexSheet, IndexPulse
        Me.Load_Pwd .LstVwPWD
    End With
    
End Sub

Sub Load_Pwd(LstVw As ListView)

Dim lv_Pls_Count            As Long
Dim lv_Index_Ini            As Long


    lv_Index_Ini = PV_Index_Pls_Ini - PV_Count_Pls / 2
    If lv_Index_Ini < 0 Then
        lv_Index_Ini = 0
    End If
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
    'LstVw.SelectedItem.
    
End Sub

Private Sub LstVwErrorList_KeyDown(KeyCode As Integer, Shift As Integer)

    Me.ShowPulseList

End Sub

Private Sub LstVwErrorList_KeyPress(KeyAscii As Integer)

    Me.ShowPulseList

End Sub

Private Sub LstVwErrorList_KeyUp(KeyCode As Integer, Shift As Integer)

    Me.ShowPulseList

End Sub

Private Sub tmrPointer_Timer()

    If moviendo = False Then
        Set_MousePointer vbDefault
    End If
    
End Sub

Sub RefreshLstVwData(LstVw As ListView, IndexFn As Integer)

    With Me
        If IndexFn = 0 Then
            .Load_Error_List LstVw
        Else
            '.Load_Pwd LstVw
        End If
    End With
    
End Sub

Sub ProcessScroll(LstVw As ListView, _
                    IndexFn As Integer, _
                    IndexCmd As Integer, _
                    IndexIni As Long, _
                    List_Count As Long)

Dim Count           As Long

    Count = ListViewGetVisibleCount(LstVw) '- 1
    With Me
        Select Case IndexCmd
            Case Is = 1
                If IndexIni Then
                    IndexIni = 0
                    RefreshLstVwData LstVw, IndexFn
                End If
            Case Is = 2
                If IndexIni < Count Then
                    If IndexIni <> 0 Then
                        IndexIni = 0
                        RefreshLstVwData LstVw, IndexFn
                    End If
                Else
                    IndexIni = IndexIni - Count
                    RefreshLstVwData LstVw, IndexFn
                End If
            Case Is = 3
                If IndexIni + Count < List_Count - 1 Then
                    IndexIni = IndexIni + Count
                    RefreshLstVwData LstVw, IndexFn
                End If
            Case Is = 4
                If IndexIni <> List_Count - Count - 1 Then
                    IndexIni = List_Count - Count - 1
                    RefreshLstVwData LstVw, IndexFn
                End If
        End Select
    End With
    
End Sub
                    
Private Sub toolbarControles_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    With Me
        If Index = 0 Then
            ProcessScroll .LstVwErrorList, Index, Button.Index, _
                            PV_Index_Err_Ini, PV_ErrList_Count
        Else
            ProcessScroll .LstVwPWD, Index, Button.Index, _
                            PV_Index_Pls_Ini, PV_PlsList_Count
        End If
    End With

End Sub

'            Select Case Button.Index
'                Case Is = 1
'                    If PV_Index_Err_Ini Then
'                        PV_Index_Err_Ini = 0
'                        .Load_Error_List .LstVwErrorList
'                    End If
'                Case Is = 2
'                    If PV_Index_Err_Ini < PV_Count_Err Then
'                        If PV_Index_Err_Ini <> 0 Then
'                            PV_Index_Err_Ini = 0
'                            .Load_Error_List .LstVwErrorList
'                        End If
'                    Else
'                        PV_Index_Err_Ini = PV_Index_Err_Ini - PV_Count_Err
'                        .Load_Error_List .LstVwErrorList
'                    End If
'                Case Is = 3
'                    If PV_Index_Err_Ini + PV_Count_Err < PV_ErrList_Count - 1 Then
'                        PV_Index_Err_Ini = PV_Index_Err_Ini + PV_Count_Err
'                        .Load_Error_List .LstVwErrorList
'                    End If
'                Case Is = 4
'                    If PV_Index_Err_Ini <> PV_ErrList_Count - PV_Count_Err - 1 Then
'                        PV_Index_Err_Ini = PV_ErrList_Count - PV_Count_Err - 1
'                        .Load_Error_List .LstVwErrorList
'                    End If
'            End Select
'        Else
'        End If
'    End With
'
'End Sub


