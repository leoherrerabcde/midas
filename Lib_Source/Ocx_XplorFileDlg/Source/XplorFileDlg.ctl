VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl XplorFileDlg 
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   Picture         =   "XplorFileDlg.ctx":0000
   ScaleHeight     =   6270
   ScaleWidth      =   6240
   Begin VB.Frame FrameSelPath 
      Caption         =   "Selección Directorio"
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin MSComctlLib.ImageCombo imgCboAbPls 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "imgCboAbPls"
      End
      Begin VB.ComboBox cboAbPls 
         Height          =   315
         Left            =   480
         TabIndex        =   14
         Text            =   "cboAbPls"
         Top             =   1560
         Width           =   4575
      End
      Begin VB.TextBox txtFileSelected 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "txtFileSelected"
         Top             =   3960
         Width           =   5895
      End
      Begin VB.CommandButton cmdOpenDatos 
         Caption         =   "cmdOpenDatos"
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   5400
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtPathPulsos 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "txtPathPulsos"
         Top             =   3960
         Width           =   5895
      End
      Begin VB.ComboBox cboExtensiones 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Text            =   "cboExtensiones"
         Top             =   4440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.PictureBox PictureBtnsFileOP 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   8520
         ScaleHeight     =   3975
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PictureBotonesDir 
         Height          =   495
         Left            =   4800
         ScaleHeight     =   435
         ScaleWidth      =   1155
         TabIndex        =   3
         Top             =   240
         Width           =   1215
         Begin VB.Image btn_vistas 
            Height          =   315
            Left            =   780
            ToolTipText     =   "Menú Ver"
            Top             =   60
            Width           =   345
         End
         Begin VB.Image btn_Nueva 
            Height          =   315
            Left            =   420
            ToolTipText     =   "Crear nueva carpeta"
            Top             =   60
            Width           =   315
         End
         Begin VB.Image btn_Nsuperior 
            Height          =   315
            Left            =   60
            ToolTipText     =   "Subir un nivel"
            Top             =   60
            Width           =   315
         End
      End
      Begin VB.Timer tmrEfecto3D 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   5400
      End
      Begin VB.TextBox txtNvaExtension 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Text            =   "txtNvaExtension"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddNvaExtension 
         Caption         =   "+"
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   4920
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSComctlLib.ListView LstVwAbrirPulsos 
         Height          =   3015
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   0
         Left            =   2880
         Top             =   5520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   1
         Left            =   3480
         Top             =   5520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label lbl 
         Caption         =   "ShComboBoxAbPls"
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
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   5400
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lbl 
         Caption         =   "FileViewAbrirPls"
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
         Left            =   840
         TabIndex        =   12
         Top             =   5760
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblExtension 
         Caption         =   "lblExtension"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblNvaExtension 
         Caption         =   "Nueva Extension:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "mnuVer"
      Visible         =   0   'False
      Begin VB.Menu mnuVistaMiniatura 
         Caption         =   "mnuVistaMiniatura"
      End
      Begin VB.Menu mnuVistaIconos 
         Caption         =   "mnuVistaIconos"
      End
      Begin VB.Menu mnuVistaLista 
         Caption         =   "mnuVistaLista"
      End
      Begin VB.Menu mnuVistaDetalles 
         Caption         =   "mnuVistaDetalles"
      End
   End
End
Attribute VB_Name = "XplorFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : OpenDlg
' Author    : lherrera
' Date      : 03/02/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_nBotonFileOp                     As Integer
Private PC_Boton3D                          As clsBoton3D
Private PC_TxtNvaExt                        As clsTextBox
Private PV_Path_File                        As String
'Public PV_Name_File                         As String
Private PV_Section                          As String
Private PV_Key                              As String
Private PV_File_Settings                    As String
Private PV_File_Selected                    As String

Private cSubLV As cSubclassListView
Private WithEvents cFile As clsListFile
Attribute cFile.VB_VarHelpID = -1

Public GV_App_Path                  As String


Public Enum OpenDialogConstant
    OpenFolder = 1
    OpenFiles
End Enum
Private PV_OpDlg_Behavior                   As OpenDialogConstant

'

Private PV_Block_Change_Path                As Boolean

Public Event PathChanged()
Public Event CmdOpenClick()
Public Event CmdCancelClick()
Public Event FileSelected()
Public Event FileNonSelected()
Public Event FileClicked(lvFile As String)
Public Event PathClicked(lvPath As String)

Property Get HandlerFileLog() As Integer

    HandlerFileLog = GV_hFile
    
End Property

Property Get Path_Iconos() As String

    Path_Iconos = GV_Path_Iconos
    
End Property

Property Get File_Btn_Nsuperior() As String

    File_Btn_Nsuperior = GV_File_Btn_Nsuperior
    
End Property

Property Let App_Path(lvPath As String)

    GV_App_Path = lvPath
    
End Property

Property Get App_Path() As String

    App_Path = GV_App_Path
    
End Property

Property Let DialogBehavior(NewBehavior As Integer)
    
    PV_OpDlg_Behavior = NewBehavior
    
End Property

Property Get DialogBehavior() As Integer
    
    DialogBehavior = PV_OpDlg_Behavior
    
End Property

Property Let FileName(l As String)
    
End Property

Property Get FileName() As String
    
    FileName = PV_File_Selected
    
End Property

Property Let Version(lv_T As String)

End Property

Property Get Version() As String
    
    Version = App.Major
    Version = Version & "." & App.Minor
    Version = Version & "." & App.Revision
    'App.RetainedProject
    
End Property

Property Let FrameWidth(lv_T As Single)

End Property

Property Get FrameWidth() As Single
    
    FrameWidth = UserControl.FrameSelPath.Width
    
End Property

Property Let FrameHeight(lv_T As Single)

End Property

Property Get FrameHeight() As Single

    FrameHeight = UserControl.FrameSelPath.Height
    
End Property

Property Let FileSettings(lvPath As String)

Dim lvValue         As String

    lvValue = lvPath
    Set_File_Settings lvValue

End Property

Property Get FileSettings() As String

    FileSettings = PV_File_Settings
    
End Property


Property Get LastPath() As String

    LastPath = PV_Path_File
    
End Property

Property Let LastPath(lvPath As String)

    PV_Path_File = lvPath
    
End Property

Property Let MultiSelectFileState(lvState As Variant)

    'UserControl.LstVwFileList.MultiSelect = lvState
    
End Property

'Property Get ControlWidth() As Variant
'
'    ControlWidth = UserControl.PictureBtnsFileOP + _
'                    UserControl.Image_Btn_File_Operation(0).left - _
'                    UserControl.FrameSelPath.left
'
'End Property
'
'Property Let LstVwSelFileState(lvState As Variant)
'
'Dim lvValue         As Boolean
'
'    lvValue = lvState
'    Set_LstVw_SelFile_State lvValue
'    Set_Btn_File_Op_State lvValue
'    UserControl.FrameSelPath.Width = UserControl.PictureBtnsFileOP + _
'                                    UserControl.Image_Btn_File_Operation(0).left - _
'                                    UserControl.FrameSelPath.left
'    UserControl.Width = UserControl.PictureBtnsFileOP + _
'                        UserControl.Image_Btn_File_Operation(0).left - _
'                        UserControl.FrameSelPath.left + _
'                        60
'
'End Property

Property Let NewExtensionState(lvState As Variant)

Dim lvValue         As Boolean

    lvValue = lvState
    Set_NewExtension_State lvValue

End Property

Property Let CommandOpenState(lvState As Variant)

Dim lvValue         As Boolean

    lvValue = lvState
    Set_CmdOpen_State lvValue

End Property

'Sub Set_Btn_File_Op_State(lvVisible As Boolean)
'
'    UserControl.PictureBtnsFileOP.Visible = lvVisible
'
'End Sub
'
'Sub Set_LstVw_SelFile_State(lvVisible As Boolean)
'
'    UserControl.LstVwFileSelected.Visible = lvVisible
'
'End Sub

Sub Set_NewExtension_State(lvVisible As Boolean)

    UserControl.lblNvaExtension.Visible = lvVisible
    UserControl.txtNvaExtension.Visible = lvVisible
    
End Sub

Sub Set_CmdOpen_State(lvVisible As Boolean)

    UserControl.cmdOpenDatos.Visible = lvVisible
    
End Sub

Sub Set_File_Settings(lvPath As String)

    PV_File_Settings = lvPath
    
End Sub

Private Sub Add_New_Filtro_Extension(ByRef lvNewExt As String, ByRef lv_CboBox As ComboBox)

Dim LV_Extension        As String
Dim Index               As Integer

    Index = 0
    With lv_CboBox
        .Clear
        Do
            LV_Extension = Leer_Ini(PV_File_Settings, _
                            .Name, _
                            .Name & ".List(" & Trim$(Index) & ")", _
                            CT_EXTENSION_DEFAULT)
            Grabar_Ini PV_File_Settings, _
                        .Name, _
                        .Name & ".List(" & Trim$(Index) & ")", _
                        LV_Extension
            .AddItem LV_Extension
            If LV_Extension = CT_EXTENSION_DEFAULT Then
                .AddItem lvNewExt
                Grabar_Ini PV_File_Settings, _
                            .Name, _
                            .Name & ".List(" & Trim$(Index) & ")", _
                            lvNewExt
                Grabar_Ini PV_File_Settings, _
                                    .Name, _
                                    .Name & ".ListIndex", _
                                    .ListCount - 1
                Exit Do
            End If
            Index = Index + 1
        Loop
        .ListIndex = Leer_Ini(PV_File_Settings, _
                            .Name, _
                            .Name & ".ListIndex", _
                            0)
        Grabar_Ini PV_File_Settings, _
                            .Name, _
                            .Name & ".ListIndex", _
                            .ListIndex
    End With
    
End Sub

Private Sub Get_Filtros_Extension(ByRef lv_CboBox As ComboBox)

Dim LV_Extension        As String
Dim Index               As Integer

    Index = 0
    On Error Resume Next
    With lv_CboBox
        .Clear
        Do
            LV_Extension = Leer_Ini(PV_File_Settings, _
                            .Name, _
                            .Name & ".List(" & Trim$(Index) & ")", _
                            CT_EXTENSION_DEFAULT)
            Grabar_Ini PV_File_Settings, _
                        .Name, _
                        .Name & ".List(" & Trim$(Index) & ")", _
                        LV_Extension
            .AddItem LV_Extension
            If LV_Extension = CT_EXTENSION_DEFAULT Then
                Exit Do
            End If
            Index = Index + 1
        Loop
        
        .ListIndex = Leer_Ini(PV_File_Settings, _
                            .Name, _
                            .Name & ".ListIndex", _
                            0)
        Grabar_Ini PV_File_Settings, _
                            .Name, _
                            .Name & ".ListIndex", _
                            .ListIndex
    End With
    On Error GoTo 0
    
End Sub

Sub Guardar_New_Path(lvNewPath As String)

    With UserControl
        If PV_Path_File <> lvNewPath Then
            'If PV_Section <> "" And PV_Key <> "" Then
                Save_Config "Configuraciones\" & PV_Section, PV_Key, lvNewPath
            'End If
            PV_Path_File = lvNewPath
        End If
    End With

End Sub

Sub Guardar_Ultimo_Path()

    With UserControl
        If PV_Section <> "" And PV_Key <> "" Then
            Save_Config "Configuraciones\" & PV_Section, PV_Key, cFile.Path
            .txtPathPulsos.Text = cFile.Path
        End If
        PV_Path_File = .txtPathPulsos.Text
    End With

End Sub

Private Sub Guardar_Vista_ListView(lvLstVw As ListView)

    If PV_Section <> "" And PV_Key <> "" Then
        Save_Config "Configuraciones\" & PV_Section, lvLstVw.Name & ".ViewStyle", lvLstVw.View
    End If
    
End Sub

Private Sub Init_Controls_Botones(LV_Path As String)

Dim lv_Path_Iconos  As String
Dim lsFile          As String
Dim i               As Integer

    With UserControl
        lv_Path_Iconos = Get_Config("Configuraciones", "lv_Path_Iconos", "Iconos")
        lv_Path_Iconos = LV_Path & "\" & lv_Path_Iconos & "\"
        
        GV_Path_Iconos = lv_Path_Iconos
        SaveSetting App.Title, "Init_Controls_Botones", "lv_Path_Iconos", lv_Path_Iconos
        'WriteLogFile "lv_path_iconos: " & lv_Path_Iconos
        lsFile = Get_Config("Configuraciones", "Icono SubirNivel", "SubirNivel")
        
        GV_File_Btn_Nsuperior = lv_Path_Iconos & lsFile & ".ico"
        'WriteLogFile "btn_Nsuperior: " & lv_Path_Iconos & lsFile & ".ico"
        btn_Nsuperior.Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
        lsFile = Get_Config("Configuraciones", "Icono CarpetaNueva", "CrearCarpeta")
        btn_Nueva.Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
        lsFile = Get_Config("Configuraciones", "Icono Vista", "Vista")
        btn_vistas.Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
'        lsFile = Get_Config("Configuraciones", "Icono Select_All", "Select_All")
'        .Image_Btn_File_Operation(0).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
'
'        lsFile = Get_Config("Configuraciones", "Icono Deselect", "Deselect")
'        .Image_Btn_File_Operation(1).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
'
'        lsFile = Get_Config("Configuraciones", "Icono Move_To", "Move_To")
'        .Image_Btn_File_Operation(2).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
'
'        lsFile = Get_Config("Configuraciones", "Icono Remove_From", "Remove_From")
'        .Image_Btn_File_Operation(3).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
'        lsFile = Get_Config("Configuraciones", "Icono Add_All", "Add_All")
'        .Image_Btn_File_Operation(0).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
'
'        lsFile = Get_Config("Configuraciones", "Icono Add_All", "Add_All")
'        .Image_Btn_File_Operation(0).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
    End With
    
End Sub

Sub Init_Ocx()

    With UserControl
        '
        Set PC_Boton3D = New clsBoton3D
        
        '
        Set PC_TxtNvaExt = New clsTextBox
        PC_TxtNvaExt.Iniciar .txtNvaExtension
        .txtNvaExtension.Text = ""
        
        .txtFileSelected.Text = ""
        
        .PictureBotonesDir.BorderStyle = 0
        
        PV_Block_Change_Path = False
        
        PV_Section = "OpenDlg"
        PV_Key = "Last Path"
        
    End With
        
End Sub

Sub Set_Open_Dialog_Behavior(lvBehavior As OpenDialogConstant)

    With UserControl
        PV_OpDlg_Behavior = lvBehavior
        '.txtFileSelected.Top = .txtPathPulsos.Top
'        If PV_OpDlg_Behavior = OpenFolder Then
'            .txtPathPulsos.Visible = True
'            .txtFileSelected.Visible = False
'        Else
'            .txtPathPulsos.Visible = False
'            .txtFileSelected.Visible = True
'        End If
    End With
    
End Sub

Sub Init_Controls()

Dim LV_Path         As String
Dim lv_Path_Iconos  As String
Dim lsFile          As String

    Set cSubLV = New cSubclassListView
    Set cFile = New clsListFile
    
    cSubLV.SubClassListView UserControl.LstVwAbrirPulsos.hwnd
    
    With UserControl
       
        If PV_OpDlg_Behavior = OpenFolder Then
            .txtPathPulsos.Visible = True
            .txtFileSelected.Visible = False
        Else
            .txtPathPulsos.Visible = False
            .txtFileSelected.Visible = True
        End If
        '
       
       ' Filtros de Extension
        Get_Filtros_Extension .cboExtensiones
        
        
        ' Cargar Íconos e Imágenes
        If GV_App_Path = "" Then
            LV_Path = Retroceder_Path(App.Path)
        Else
            LV_Path = Retroceder_Path(GV_App_Path)
        End If
        
        SaveSetting App.Title, "app", ".Path", App.Path
        Init_Controls_Botones LV_Path
        
        LV_Path = Obtener_Ultimo_Path(LV_Path)
        
        cFile.Path = LV_Path
        Recuperar_Vista_LstVw .LstVwAbrirPulsos
        
        cFile.SetControls .LstVwAbrirPulsos, ImageList1(0), ImageList1(1)
        cFile.Listar LV_Path
        
        Set_Open_Dialog_Behavior PV_OpDlg_Behavior
        
    End With

End Sub

Function Obtener_Ultimo_Path(lsDefault As String) As String

    Obtener_Ultimo_Path = PV_Path_File
'    If PV_Section <> "" And PV_Key <> "" Then
'        Obtener_Ultimo_Path = Get_Config("Configuraciones", "Ultimo Path", lsDefault)
'    End If
    
End Function

Private Sub Put_Files_On_List(LV_Path As String, lv_Ext As String, lvLstVw As ListView)

Dim lv_File_List()          As String
Dim LV_Count                As Integer
Dim LV_Nom_Arch             As String

    LV_Count = 0
    ReDim lv_File_List(LV_Count + 100)
    
    
    LV_Nom_Arch = Dir(LV_Path & "\" & lv_Ext, vbNormal)
    Do While Len(LV_Nom_Arch) > 0
        If (LV_Nom_Arch <> "..") And (LV_Nom_Arch <> ".") Then
            lv_File_List(LV_Count) = LV_Nom_Arch
            LV_Count = LV_Count + 1
            If UBound(lv_File_List) < LV_Count Then
                ReDim Preserve lv_File_List(LV_Count + 100)
            End If
        Else
        End If
        LV_Nom_Arch = Dir()
    Loop
    
'    If LV_Count Then
'        ReDim Preserve lv_File_List(LV_Count - 1)
'        LstVw_Add_ListItems lvLstVw, lv_File_List, True
'        LVSetAllColWidths lvLstVw, LVSCW_AUTOSIZE
'    End If
    
End Sub


Private Sub Recuperar_Vista_LstVw(lvFileVw As ListView)

    If PV_Section <> "" And PV_Key <> "" Then
        lvFileVw.View = Get_Config("Configuraciones\" & PV_Section, lvFileVw.Name & ".ViewStyle", lvFileVw.View)
    End If
    
End Sub

Private Sub btn_Nsuperior_Click()

    cFile.subirNivel
    
End Sub

Private Sub btn_Nsuperior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With UserControl
        PC_Boton3D.Iniciar .btn_Nsuperior, .PictureBotonesDir, .tmrEfecto3D
    End With
    
End Sub

Private Sub btn_Nueva_Click()

Dim lvPath          As String

    frmNuevaCarpeta.Set_Last_Path UserControl.txtPathPulsos.Text
    frmNuevaCarpeta.Show vbModal
    lvPath = cFile.Path
    cFile.Path = lvPath
    
End Sub

Private Sub btn_Nueva_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With UserControl
        PC_Boton3D.Iniciar UserControl.btn_Nueva, UserControl.PictureBotonesDir, UserControl.tmrEfecto3D
    End With

End Sub

Private Sub btn_vistas_Click()

    With UserControl
        .PopupMenu .mnuVer
    End With
    
End Sub

Private Sub btn_vistas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With UserControl
        PC_Boton3D.Iniciar .btn_vistas, .PictureBotonesDir, .tmrEfecto3D
    End With

End Sub

Private Sub cboExtensiones_Change()

    With UserControl.cboExtensiones
        Grabar_Ini PV_File_Settings, _
                            .Name, _
                            .Name & ".ListIndex", _
                            .ListIndex
    End With
    
End Sub

Private Sub Save_Config(lsSection As String, lsKey As String, lValue As Variant)

    Grabar_Ini PV_File_Settings, _
                        lsSection, _
                        lsKey, _
                        lValue

End Sub

Private Function Get_Config(lsSection As String, lsKey As String, lDefault As Variant) As String

    Get_Config = Leer_Ini(PV_File_Settings, _
                        lsSection, _
                        lsKey, _
                        lDefault)
    
End Function

Private Sub cFile_changePath(Ruta As String)

    If PV_Block_Change_Path = False Then
        PV_Block_Change_Path = True
        If cFile.Path <> "" Then
            New_Path_Selected cFile.Path
        End If
        PV_Block_Change_Path = False
    End If

End Sub

Private Sub cmdAddNvaExtension_Click()

    With UserControl
        Add_New_Filtro_Extension .txtNvaExtension.Text, .cboExtensiones
    End With
    
End Sub

Private Sub cmdOpenDatos_Click()

    With UserControl
'        Open_File fMainForm.CommonDialog, GV_Path, lvFilter, "Abrir Datos..."
'        BrowseForFolder UserControl.hWnd, "Seleccionar Directorio"
        Guardar_Ultimo_Path
    End With
    
    'Unload Me
    
End Sub



Sub Raise_Path_Selected()

Dim lsPath          As String

    On Error Resume Next
    With UserControl
        lsPath = .LstVwAbrirPulsos.SelectedItem.Text
        If lsPath <> "" Then
            lsPath = cFile.Path & "\" & lsPath
        Else
            lsPath = cFile.Path
        End If
        If Is_Folder(lsPath) = True Then
            RaiseEvent PathClicked(lsPath)
        Else
            RaiseEvent FileClicked(lsPath)
        End If
    End With
    
End Sub

Private Sub New_Path_Selected(lvNewPath As String)

    Guardar_New_Path lvNewPath
    
    If UserControl.txtPathPulsos.Text <> lvNewPath Then
        UserControl.txtPathPulsos.Text = lvNewPath
    End If
'    If UserControl.ShComboBoxAbPls.SelectedItem.Path <> lvNewPath Then
'        UserControl.ShComboBoxAbPls.SetPath lvNewPath
'    End If
    If UserControl.LstVwAbrirPulsos.SelectedItem.Path <> lvNewPath Then
        UserControl.LstVwAbrirPulsos.SetPath lvNewPath
    End If
    If cFile.Path <> lvNewPath Then
        cFile.Path = lvNewPath
    End If
    Put_Files_On_List UserControl.txtPathPulsos.Text, _
                        UserControl.cboExtensiones.Text ', _
                        'UserControl.LstVwFileList
    RaiseEvent PathChanged

End Sub

' True is a valid file
Private Function Set_New_Selected_File(lv_NewFile As String) As Boolean

Dim lvFile           As String
    
    Set_New_Selected_File = False
    
    On Error GoTo Isn_a_File
    
    With UserControl
        lvFile = PV_Path_File & "\" & lv_NewFile
        If GetAttr(lvFile) = vbArchive Then
            .txtFileSelected.Text = lv_NewFile
            PV_File_Selected = lv_NewFile
            Set_New_Selected_File = True
            RaiseEvent FileSelected
        Else
            RaiseEvent FileNonSelected
        End If
    End With
    
    On Error GoTo 0
    
    Exit Function
    
Isn_a_File:
    On Error GoTo 0
    RaiseEvent FileNonSelected
    
End Function


Private Sub Form_Load()

    'Init_Controls
    
End Sub

'Private Sub Image_Btn_File_Operation_Click(Index As Integer)
'
'    'Si es el botón izquierdo...
'
'    With UserControl
''        If Button = 1 Then
''            'Efecto de pulsado
''            .Image_Btn_File_Operation(Index).BorderStyle = 1
''            'lblTip.Visible = False
''            'cargar la rutina correspondiente
''            Image_Btn_File_Operation(Index).BorderStyle = 0
''        End If
'        Select Case Index
'            Case 0
'                ' Select All
'                'LstVw_Set_CheckBox .LstVwFileList, True
'            Case 1
'                ' Deselect All
'                'LstVw_Set_CheckBox .LstVwFileList, False
'            Case 2
'                ' Move To
'                'LstVw_Move_Chequed_To .LstVwFileList, .LstVwFileSelected
'            Case 3
'                ' Remove From
'                'LstVw_Move_Selected_To .LstVwFileSelected, .LstVwFileList
'        End Select
'    End With
'
'End Sub

'Private Sub Image_Btn_File_Operation_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    With UserControl
'        PC_Boton3D.Iniciar .Image_Btn_File_Operation(Index), _
'                            .PictureBtnsFileOP, _
'                            .tmrEfecto3D
'    End With
'
'End Sub

'Private Sub LstVwFileList_DblClick()
'
'    LstVw_Check_Selected UserControl.LstVwFileList, True
'    RaiseEvent FileSelected
'
'End Sub

Private Sub LstVwAbrirPulsos_Click()

    With UserControl
        If .LstVwAbrirPulsos.SelectedItem.Index <= 0 Then
            .cmdOpenDatos.Enabled = False
        Else
                If PV_OpDlg_Behavior = OpenFiles Then
                    If Set_New_Selected_File(.LstVwAbrirPulsos.SelectedItem.Text) = True Then
                        .cmdOpenDatos.Enabled = True
                        .txtPathPulsos = .txtFileSelected
                    End If
                Else
                    Raise_Path_Selected
                End If
        End If
        
    End With
    
End Sub

Private Sub LstVwAbrirPulsos_DblClick()

    If PV_OpDlg_Behavior = OpenFiles Then
        RaiseEvent CmdOpenClick
    End If
    
End Sub

Private Sub LstVwAbrirPulsos_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Set_New_Selected_File Item.Text

End Sub

Private Sub LstVwAbrirPulsos_KeyDown(KeyCode As Integer, Shift As Integer)

    Raise_Path_Selected
    
End Sub

Private Sub LstVwAbrirPulsos_KeyPress(KeyAscii As Integer)

    Raise_Path_Selected
    
End Sub

Private Sub LstVwAbrirPulsos_KeyUp(KeyCode As Integer, Shift As Integer)

    Raise_Path_Selected
    
End Sub

Private Sub mnuVistaDetalles_Click()

    With UserControl
        .LstVwAbrirPulsos.View = lvwReport
        Guardar_Vista_ListView .LstVwAbrirPulsos
    End With
    
End Sub

Private Sub mnuVistaIconos_Click()

    With UserControl
        .LstVwAbrirPulsos.View = lvwIcon
        Guardar_Vista_ListView .LstVwAbrirPulsos
    End With
        
End Sub

Private Sub mnuVistaLista_Click()

    With UserControl
        .LstVwAbrirPulsos.View = lvwList
        Guardar_Vista_ListView .LstVwAbrirPulsos
    End With
    
End Sub

Private Sub mnuVistaMiniatura_Click()

    With UserControl
        .LstVwAbrirPulsos.View = lvwSmallIcon
        Guardar_Vista_ListView .LstVwAbrirPulsos
    End With
    
End Sub

Private Sub txtFileSelected_Change()

    PV_File_Selected = UserControl.txtFileSelected
    
End Sub

Private Sub txtPathPulsos_Change()

    With UserControl
        '.txtFileSelected = .txtPathPulsos
    End With

End Sub

Private Sub UserControl_Initialize()

    Call InitCommonControls

    Call SetErrorMode(2)
    
    Set_Open_Dialog_Behavior OpenFiles
    
    Init_Ocx

End Sub

Private Sub UserControl_Resize()

'Dim lWidth          As Long

    On Error Resume Next
    'WriteLogFile "app.Path = " & App.Path
    With UserControl
        .FrameSelPath.Height = .Height - 45
        '.LstVwFileList.Height = .FrameSelPath.Height - 45 - .LstVwFileList.Top
        .FrameSelPath.Width = .Width - 60
        'lWidth = .FrameSelPath.Width - .LstVwFileList.left - 50
        'If lWidth >= 1200 Then
        '    .LstVwFileList.Width = lWidth
        'End If
        
        .LstVwAbrirPulsos.Width = .FrameSelPath.Width - _
                                    2 * .LstVwAbrirPulsos.left
        .PictureBotonesDir.left = .LstVwAbrirPulsos.left + _
                                    .LstVwAbrirPulsos.Width - _
                                    .PictureBotonesDir.Width
        .cboAbPls.Width = .PictureBotonesDir.left - 2 * .cboAbPls.left
        .txtFileSelected.Top = .FrameSelPath.Height - .txtFileSelected.Height - 120
        .txtFileSelected.Width = .LstVwAbrirPulsos.Width
        .LstVwAbrirPulsos.Height = .txtFileSelected.Top - .LstVwAbrirPulsos.Top - 120
        .txtPathPulsos.Top = .txtFileSelected.Top
        .txtPathPulsos.Width = .txtFileSelected.Width
        
    End With
    On Error GoTo 0
        
End Sub

Private Sub UserControl_Terminate()

    CloseLogFile
    
End Sub
