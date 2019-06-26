VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPjtFromList 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   9045
   Begin VB.Frame FrameProjectRecentList 
      Caption         =   "Proyectos Recientes"
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      Begin VB.CommandButton cmdAccept 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6960
         TabIndex        =   3
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   3960
         Width           =   975
      End
      Begin MSComctlLib.ListView LstVwRecentProjectList 
         Height          =   3495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6165
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Indice"
            Object.Width           =   953
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ubicación"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPjtFromList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmPjtFromList
' Author    : Leo Herrera
' Date      : 05/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Sub Init_Form()

Dim LV_Path             As String

    On Error Resume Next
    Me.Caption = "Abrir Proyecto Reciente"
    Me.WindowState = vbMaximized
    LoadColumnWideListView Me.LstVwRecentProjectList, App.Title, Me.Name
    LoadListRecentsProjects Me.LstVwRecentProjectList

End Sub

Sub LoadListRecentsProjects(LV_LstVw As ListView)

    m_Project.ShowListProjects LV_LstVw

End Sub

Sub LoadSelectedProject()

Dim lvProject           As String
Dim lvPathProject       As String

    With Me.LstVwRecentProjectList
        lvProject = .SelectedItem.ListSubItems(1).Text
        lvPathProject = .SelectedItem.ListSubItems(2).Text
        
        m_Project.LoadProject lvProject, lvPathProject
        
        GV_Project_Opened = True
        
        Unload Me
    End With
    ' Only for Debugging
    If GV_Debug.Template = True Then
        GV_clsTemplateConfigSpreadSheet.SetLoad
    End If
    
End Sub

Sub OpenProjectFromListView(LV_LstView As ListView)

End Sub

Private Sub cmdAccept_Click()

    LoadSelectedProject
    
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    GV_Project_Closed = True
    
End Sub

Private Sub Form_Load()

    Init_Form
    Set_MousePointer vbDefault
    
End Sub

Private Sub Form_Resize()

Dim LV_Width                As Long
Dim i                       As Integer
Dim lvMargen                As Long

    lvMargen = 120
    With Me
        .FrameProjectRecentList.Width = .ScaleWidth - _
                                        2 * .FrameProjectRecentList.Left
        .LstVwRecentProjectList.Width = .FrameProjectRecentList.Width - _
                                        2 * .LstVwRecentProjectList.Left
        With .LstVwRecentProjectList
            LV_Width = .Width
            For i = 1 To .ColumnHeaders.Count - 1
                LV_Width = LV_Width - .ColumnHeaders.Item(i).Width
            Next
            .ColumnHeaders(.ColumnHeaders.Count).Width = LV_Width
        End With
        
        .FrameProjectRecentList.Height = .ScaleHeight - _
                                        2 * .FrameProjectRecentList.Top
        .cmdAccept.Top = .FrameProjectRecentList.Height - .cmdAccept.Height - _
                        lvMargen
        .cmdCancel.Top = .cmdAccept.Top
        .LstVwRecentProjectList.Height = .cmdAccept.Top - _
                                         .LstVwRecentProjectList.Top - _
                                         lvMargen
        
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveColumnWideListView Me.LstVwRecentProjectList, App.Title, Me.Name
    GV_clsPjtFromList.ClearLoaded
    
End Sub

Private Sub LstVwRecentProjectList_DblClick()

    LoadSelectedProject

End Sub

Private Sub LstVwRecentProjectList_ItemClick(ByVal Item As MSComctlLib.ListItem)

    'LoadSelectedProject
    
End Sub

Private Sub LstVwRecentProjectList_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc(vbCr) Then
        LoadSelectedProject
    End If
    
End Sub
