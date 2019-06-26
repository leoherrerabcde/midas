VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNavProject 
   BorderStyle     =   0  'None
   Caption         =   "0"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView LstVwGralInfo 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox PictureImBtn 
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   2295
      TabIndex        =   6
      Top             =   3720
      Width           =   2355
      Begin VB.Image ImageBtnDatos 
         Height          =   435
         Index           =   0
         Left            =   60
         Stretch         =   -1  'True
         Top             =   60
         Width           =   435
      End
      Begin VB.Image ImageBtnDatos 
         Height          =   435
         Index           =   1
         Left            =   600
         Stretch         =   -1  'True
         Top             =   60
         Width           =   435
      End
      Begin VB.Image ImageBtnDatos 
         Height          =   435
         Index           =   2
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   60
         Width           =   435
      End
      Begin VB.Image ImageBtnDatos 
         Height          =   435
         Index           =   3
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   60
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdSelectFunction 
      Caption         =   "cmdAssignData"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   2355
   End
   Begin VB.CommandButton cmdSelectFunction 
      Caption         =   "cmdStructData"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   2355
   End
   Begin VB.CommandButton cmdSelectFunction 
      Caption         =   "cmdAnalisisData"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   2355
   End
   Begin VB.CommandButton cmdSelectFunction 
      Caption         =   "cmdGralInfo"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   2355
   End
   Begin VB.PictureBox PictureBorderForm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdCerraForm 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ImageListNavPjt 
      Left            =   4080
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView TrVwListaDatos 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   4048
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmNavProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmNavProject
' Author    : lherrera
' Date      : 02/03/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_Form_PjtFileList         As frmPjtFileList
Private PV_Set_Size_Pending()       As Boolean

Enum GralInfoControlConstants
    giccListView = 1
    giccTreeView
End Enum

Public Sub Init_Form()

    ' Config Image to works as Buttoms
    Me.PictureImBtn.BorderStyle = 0
    Set PV_Form_PjtFileList = Nothing
    
    ' Adjust the Form's Size
    Me.Show
    
    ' Seleccionar Vista por Defecto
    Me.cmdSelectFunction(0).Value = True
    
    'ShowGralInfoOnFlexGrid GV_Project, Me.MSHFlexGrid
    'ReDim PV_Set_Size_Pending(Me.SSTab.Tabs)

    'Put_Data_ListOnTrVw GV_Project, Me.TrVwListaDatos

End Sub

Private Sub cmdSelectFunction_Click(Index As Integer)

    Select Case Index
        Case 0
            Load_Prj_Gral_Info
        Case 1
            GV_MDI.mnuAnalisData_Click
        Case 2
            GV_MDI.mnuStructData_Click
        Case 3
            GV_MDI.mnuAssignData_Click
    End Select

End Sub

Private Sub Show_Controls(lvOperation As GralInfoControlConstants)

    With Me
        .LstVwGralInfo.Visible = False
        .TrVwListaDatos.Visible = False
        Select Case lvOperation
            Case giccListView
                .LstVwGralInfo.Visible = True
            Case giccTreeView
                .TrVwListaDatos.Visible = True
        End Select
    End With
    
End Sub

Public Sub Load_Prj_Gral_Info()

    With Me
        ShowGralInfoAtLstVw GV_Project, Me.LstVwGralInfo
        Show_Controls giccListView
    End With
    GV_MDI.Cargar_PjtGralInfo
    
End Sub

'Private Sub Set_Control_Tab_Size(lvTab As Integer)
'
'    With Me
'        Select Case lvTab
'            Case 0
'                If PV_Set_Size_Pending(lvTab) = True Then
'                    Set_Control_Size .MSHFlexGrid, .SSTab.Width, .SSTab.Height - .SSTab.TabHeight
'                End If
'            Case 1
'                If PV_Set_Size_Pending(lvTab) = True Then
'                    Set_Control_Size .TrVwListaDatos, .SSTab.Width, .SSTab.Height - .SSTab.TabHeight
'                End If
'            Case 2
'            Case 3
'        End Select
'        PV_Set_Size_Pending(lvTab) = False
'    End With
'
'End Sub

Private Sub Set_Control_Tab_Pending()

Dim i               As Integer

    For i = 0 To UBound(PV_Set_Size_Pending)
        PV_Set_Size_Pending(i) = True
    Next
    
End Sub

Private Sub Form_Resize()

Dim i               As Integer
Static lvFlag       As Boolean

    If lvFlag = True Then
        Exit Sub
    End If
    With Me
        lvFlag = True
        Resize_Control .PictureBorderForm, Me, refoHorizontal Or refoToBorder, 60
        Move_Control_Inside_Form .cmdSelectFunction(0), Me, refoBottom, 60
        For i = 0 To Me.cmdSelectFunction.UBound
            Resize_Control Me.cmdSelectFunction(i), Me, refoHorizontal Or refoToBorder, 60
        Next
        For i = .cmdSelectFunction.LBound + 1 To .cmdSelectFunction.UBound
            Move_Control_NextTo .cmdSelectFunction(i), _
                .cmdSelectFunction(i - 1), _
                refoTop Or refoToBorder _
                , 60
        Next
        Move_Control_NextTo .PictureImBtn _
            , .cmdSelectFunction(.cmdSelectFunction.UBound) _
            , refoTop Or refoToBorder _
            , 60
        Move_Control_NextTo .LstVwGralInfo, .PictureBorderForm, refoBottom Or refoToBorder, 60
        Move_Control_NextTo .TrVwListaDatos, .PictureBorderForm, refoBottom Or refoToBorder, 60
        Resize_Control Me.PictureImBtn, Me, refoHorizontal Or refoToBorder
        Resize_Control_UpTo .LstVwGralInfo, .PictureImBtn, refoVertical Or refoToBorder, 30
        Resize_Control_UpTo .TrVwListaDatos, .PictureImBtn, refoVertical Or refoToBorder, 30
        Resize_Control .TrVwListaDatos, Me, refoHorizontal Or refoToBorder, 60
        Resize_Control .TrVwListaDatos, Me, refoHorizontal Or refoToBorder, 60
        lvFlag = False
    End With
    
End Sub

Sub Show_List_File_NodSelected()

End Sub

Private Sub ImageBtnDatos_Click(Index As Integer)

    Select Case Index
        Case 0
        ' Asignar Archivos de Datos
        Case 1
        ' Asignar Filtro de Lectura de Datos
        Case 2
        '
        Case 3
        '
    End Select
    
End Sub

'Private Sub SSTab_Click(PreviousTab As Integer)
'
'    Set_Control_Tab_Size Me.SSTab.Tab
'    Select Case Me.SSTab.Tab
'        Case 0
'        Case 1
'            If PV_Form_PjtFileList Is Nothing Then
'                If (Me.TrVwListaDatos.SelectedItem Is Nothing) = False Then
'                    Set PV_Form_PjtFileList = GV_MDI.Cargar_Form_PjtFileList
'                    PV_Form_PjtFileList.Show_Pjt_List_File Me.TrVwListaDatos.SelectedItem.Tag
'                    GV_MDI.Set_Pos_Form_Beside PV_Form_PjtFileList, Me.Name, PosRight
'                End If
'            Else
'                GV_MDI.Show_Form_Given_Name PV_Form_PjtFileList.Name
'            End If
'        Case 2
'        Case 3
'    End Select
'
'End Sub

Private Sub TrVwListaDatos_Click()

    Show_List_File_NodSelected
    
End Sub
