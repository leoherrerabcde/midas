VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPjtFromFile 
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   11715
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame FrameProjectPreview 
      Caption         =   "Información Proyecto"
      Height          =   4335
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin MSComctlLib.ListView LstVwProjectInfo 
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   6588
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox CtlOpenDlgProject 
      Height          =   4695
      Left            =   120
      ScaleHeight     =   4635
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmPjtFromFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmPjtFromFile
' Author    : Leo Herrera
' Date      : 05/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

' Ventana de Explorador
' Preview del Proyecto
' Abrir Proyecto

Sub Init_Form()

Dim LV_Path             As String

    On Error Resume Next
    Me.WindowState = vbMaximized
    Me.LstVwProjectInfo.ListItems.Clear
    Me.Caption = "Buscar un Proyecto en Disco"
    'Open_Buttom_State False
    
    With Me.CtlOpenDlgProject
        LV_Path = GetSetting(App.Title, Me.Name, .Name & ".LastPath", App.Path)
        .LastPath = LV_Path
        .Set_Btn_File_Op_State False
        .Set_CmdOpen_State False
        .Set_File_Settings Find_Folder("Config", App.Path) & "\OpenPrjCfg.txt"
        .App_Path = App.Path
        .Set_LstVw_SelFile_State False
        .Set_NewExtension_State False
        .Set_Open_Dialog_Behavior OpenFiles
        
        SaveSetting App.Title, .Name, "Set_File_Settings", Find_Folder("Config", App.Path) & "\OpenPrjCfg.txt"
        .Init_Controls
    End With
    'On Error GoTo 0
    
End Sub

Sub Init_Info_Project_LstVw(LV_LstVw As ListView)

    With LV_LstVw.ListItems
        .Clear
        .Add , , "Name"
        .Add , , "File"
        .Add , , "Ubicacion"
        .Add , , "Ubicación Pulsos"
        .Add , , "Ubicación Salida"
        .Add , , "Analisis de Pulsos"
        .Add , , "Temporales Generados"
        .Add , , "Salida Excel Generada"
        '.Add , , ""
    End With
    
End Sub

Sub ShowProjectInfo()

Dim lsProject           As String

    With Me
        lsProject = .CtlOpenDlgProject.LastPath
        lsProject = lsProject & "\" & .CtlOpenDlgProject.FileName
        m_Project.LoadProject lsProject
        m_Project.LoadInfoProject .LstVwProjectInfo
    End With
    
End Sub

Private Sub cmdAccept_Click()

    GV_Project_Opened = True
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()

    m_Project.DiscardProject
    GV_Project_Closed = True
    Unload Me
    
End Sub

Private Sub CtlOpenDlgProject_FileClicked(lvFile As String)

    ShowProjectInfo
    
End Sub

Private Sub CtlOpenDlgProject_FileSelected()

    ShowProjectInfo
    
End Sub

Private Sub Form_Load()

    Me.Init_Form
    Me.Init_Info_Project_LstVw Me.LstVwProjectInfo
    Set_MousePointer vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GV_clsPjtFromFile.ClearLoaded
    
End Sub
