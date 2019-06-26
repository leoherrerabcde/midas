VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpreadSheetView 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10230
   Icon            =   "frmSpreadSheetView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   10230
   Begin VB.Timer TimerStartUp 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   5400
   End
   Begin MSComctlLib.ListView LstVwPwd 
      Height          =   735
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1296
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBarMSHFlex 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView LstVwPulseFile 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Files"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmSpreadSheetView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmSpreadSheetView
' Author    : Leo Herrera
' Date      : 16/07/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private PV_Index_File_Issued        As Integer

Sub ShowPWDSelectedinLstVw(LstView As ListView, LstVwPWD As ListView)

    Dim Count           As Long
    Dim Index           As Integer
    Dim i, j, q         As Long
    Dim ldArray()       As Double
    Dim lCols           As Integer
    Dim CountDisplay    As Integer
    Dim MousePointerOld As Long
    Dim LstSubItms      As ListSubItems
    
    CountDisplay = 100
    Index = LstView.SelectedItem.Index - 1
    If Index + 1 = PV_Index_File_Issued Then
        Exit Sub
    Else
        PV_Index_File_Issued = Index + 1
    End If
    lCols = Pulse_Field_Count
    Count = Pulse_Count(Index)
    ReDim ldArray(lCols - 1)
    Me.LstVwPulseFile.Refresh
    Me.Enabled = False
    MousePointerOld = Me.MousePointer
    Me.MousePointer = vbHourglass
    With Me.ProgressBarMSHFlex
        .Min = 0
        .Value = 0
        .Max = Count
        .Visible = True
    End With
    With LstVwPWD
        .Visible = False
        .ListItems.Clear
        Me.Refresh
        q = 0
        For i = 0 To Count - 1
            Pulse_Get_Pwd Index, i, ldArray(0)
            AddDoubleItemListView Me.LstVwPWD, ldArray, i, True
            If q >= CountDisplay Then
                q = 0
                '.Refresh
                Me.ProgressBarMSHFlex.Refresh
            Else
                q = q + 1
            End If
            Me.ProgressBarMSHFlex.Value = i
        Next
        .Visible = True
    End With
    AutoAjusteColumnWidth Me.LstVwPWD
    Me.ProgressBarMSHFlex.Visible = False
    Me.MousePointer = MousePointerOld
    Me.Enabled = True

End Sub

Sub Init_Form()

    Dim i           As Integer
    
    With Me
        .WindowState = vbMaximized
        PV_Index_File_Issued = -1
    End With
    
End Sub

Sub Load_Pulse_File(LstView As ListView)

Dim i       As Integer
Dim lsStr   As String

    lsStr = Space(261)
    With LstView
        .ListItems.Clear
        
        For i = 0 To Pulse_Files_Count - 1
            'Pulse_Get_File i, lsStr
            .ListItems.Add , , lsStr
            'structdatos.pulse_get_file
        Next
        If .ListItems.Count Then
            .ListItems.Item(1).Selected = True
            .Refresh
        End If
    End With
    AutoAjusteColumnWidth Me.LstVwPulseFile
    
End Sub

Sub Load_Pwd_ColumnHeader(LstVw As ListView)

Dim lsStr       As String
Dim i           As Integer

    lsStr = Space(261)
    With LstVw.ColumnHeaders
        .Clear
        .Add
        For i = 0 To Pulse_Field_Count - 1
            Pulse_Field_Header i, lsStr
             .Add , , lsStr
        Next
    End With
    
End Sub

Private Sub Form_Load()

    Init_Form
    Load_Pwd_ColumnHeader Me.LstVwPWD
    Load_Pulse_File Me.LstVwPulseFile
    Me.TimerStartUp.Enabled = True
    Set_MousePointer vbDefault

End Sub

Private Sub Form_Resize()

    With Me
        .LstVwPulseFile.Height = .ScaleHeight - 2 * .LstVwPulseFile.Top
        .LstVwPWD.Width = .ScaleWidth - .LstVwPWD.Left
        .LstVwPWD.Height = .LstVwPulseFile.Height
        .ProgressBarMSHFlex.Left = .LstVwPWD.Left
        .ProgressBarMSHFlex.Width = .LstVwPWD.Width
        .ProgressBarMSHFlex.Top = .LstVwPWD.Top + _
                        .LstVwPWD.Height - _
                        .ProgressBarMSHFlex.Height
        .ProgressBarMSHFlex.Visible = False
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GV_clsConfigSpreadSheet.ClearLoaded
    'GV_Mdi.Restore_Visible_MnuProject
    
End Sub

Private Sub LstVwPulseFile_Click()

    ShowPWDSelectedinLstVw Me.LstVwPulseFile, Me.LstVwPWD
    
End Sub

Private Sub LstVwPulseFile_DblClick()

    ShowPWDSelectedinLstVw Me.LstVwPulseFile, Me.LstVwPWD
    
End Sub

Private Sub LstVwPulseFile_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    ShowPWDSelectedinLstVw Me.LstVwPulseFile, Me.LstVwPWD

End Sub

Private Sub LstVwPulseFile_ItemClick(ByVal Item As MSComctlLib.ListItem)

    ShowPWDSelectedinLstVw Me.LstVwPulseFile, Me.LstVwPWD

End Sub

Private Sub LstVwPulseFile_KeyDown(KeyCode As Integer, Shift As Integer)

    ShowPWDSelectedinLstVw Me.LstVwPulseFile, Me.LstVwPWD
    
End Sub

Private Sub LstVwPulseFile_KeyUp(KeyCode As Integer, Shift As Integer)

    ShowPWDSelectedinLstVw Me.LstVwPulseFile, Me.LstVwPWD
    
End Sub

Private Sub TimerStartUp_Timer()

    With Me
        .TimerStartUp.Enabled = False
        ShowPWDSelectedinLstVw .LstVwPulseFile, .LstVwPWD
    End With
    
End Sub
