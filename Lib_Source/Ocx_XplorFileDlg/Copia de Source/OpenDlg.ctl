VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0502C911-34E6-4C70-8983-95DC0AB6FD7A}#7.0#0"; "shcmb70.ocx"
Object = "{9395F630-158C-4120-935E-8A7F74ACE62E}#7.0#0"; "filevw70.ocx"
Begin VB.UserControl OpenDlg 
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11715
   Picture         =   "OpenDlg.ctx":0000
   ScaleHeight     =   6270
   ScaleWidth      =   11715
   Begin VB.Frame FrameSelPath 
      Caption         =   "Selección Directorio"
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdOpenDatos 
         Caption         =   "cmdOpenDatos"
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtPathPulsos 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "txtPathPulsos"
         Top             =   3960
         Width           =   5655
      End
      Begin VB.ComboBox cboExtensiones 
         Height          =   315
         Left            =   1560
         TabIndex        =   5
         Text            =   "cboExtensiones"
         Top             =   4440
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
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   7
            Left            =   120
            Top             =   3480
            Width           =   345
         End
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   6
            Left            =   120
            Top             =   3000
            Width           =   345
         End
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   5
            Left            =   120
            Top             =   2520
            Width           =   345
         End
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   4
            Left            =   120
            Top             =   2040
            Width           =   345
         End
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   3
            Left            =   120
            Top             =   1560
            Width           =   345
         End
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   2
            Left            =   120
            Top             =   1080
            Width           =   345
         End
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   1
            Left            =   120
            Top             =   600
            Width           =   345
         End
         Begin VB.Image Image_Btn_File_Operation 
            Height          =   315
            Index           =   0
            Left            =   120
            Top             =   120
            Width           =   345
         End
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
         Left            =   1560
         TabIndex        =   2
         Text            =   "txtNvaExtension"
         Top             =   4920
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddNvaExtension 
         Caption         =   "+"
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   4920
         Width           =   255
      End
      Begin MSComctlLib.ListView LstVwFileSelected 
         Height          =   5415
         Left            =   9120
         TabIndex        =   7
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   9551
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
      Begin MSComctlLib.ListView LstVwFileList 
         Height          =   5415
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
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
      Begin ShComboBox.ShComboBox ShComboBoxAbPls 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   4455
         _Version        =   458752
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FileViewControl.FileView FileViewAbrirPls 
         Height          =   3015
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   5655
         _Version        =   458752
         _ExtentX        =   9975
         _ExtentY        =   5318
         _StockProps     =   64
         AllowFileExecute=   0   'False
         AllowItemRenaming=   0   'False
         CurrentFolder   =   "OpenDlg.ctx":1B692
         AllowDragDrop   =   0   'False
         AllowZipFolders =   0   'False
         HideSelection   =   0   'False
      End
      Begin VB.Label lblExtension 
         Caption         =   "lblExtension"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label lblNvaExtension 
         Caption         =   "Nueva Extension:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4920
         Width           =   1335
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "OpenDlg"
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
'

Public Event PathChanged()
Public Event CmdOpenClick()
Public Event CmdCancelClick()
Public Event FileSelected()

Property Get Version()
    
    Version = App.Major
    Version = Version & "." & App.Minor
    Version = Version & "." & App.Revision
    
End Property

Property Let FileSettings(lvPath As Variant)

Dim lvValue         As String

    lvValue = lvPath
    Set_File_Settings lvValue

End Property

Property Get LastPath() As Variant

    LastPath = PV_Path_File
    
End Property

Property Let LastPath(lvPath As Variant)

    PV_Path_File = lvPath
    
End Property

Property Let MultiSelectFileState(lvState As Variant)

    UserControl.LstVwFileList.MultiSelect = lvState
    
End Property

Property Get ControlWidth() As Variant

    ControlWidth = UserControl.PictureBtnsFileOP + _
                    UserControl.Image_Btn_File_Operation(0).Left - _
                    UserControl.FrameSelPath.Left
    
End Property

Property Let LstVwSelFileState(lvState As Variant)

Dim lvValue         As Boolean

    lvValue = lvState
    Set_LstVw_SelFile_State lvValue
    Set_Btn_File_Op_State lvValue
    UserControl.FrameSelPath.Width = UserControl.PictureBtnsFileOP + _
                                    UserControl.Image_Btn_File_Operation(0).Left - _
                                    UserControl.FrameSelPath.Left
    UserControl.Width = UserControl.PictureBtnsFileOP + _
                        UserControl.Image_Btn_File_Operation(0).Left - _
                        UserControl.FrameSelPath.Left + _
                        60
    
End Property

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

Sub Set_Btn_File_Op_State(lvVisible As Boolean)

    UserControl.PictureBtnsFileOP.Visible = lvVisible
    
End Sub

Sub Set_LstVw_SelFile_State(lvVisible As Boolean)

    UserControl.LstVwFileSelected.Visible = lvVisible
    
End Sub

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
    
End Sub

Sub Guardar_Ultimo_Path()

    With UserControl
        If PV_Section <> "" And PV_Key <> "" Then
            Save_Config "Configuraciones\" & PV_Section, PV_Key, .FileViewAbrirPls.CurrentFolder
            .txtPathPulsos.Text = .FileViewAbrirPls.CurrentFolder
            PV_Path_File = .txtPathPulsos.Text
        End If
    End With

End Sub


Private Sub Guardar_Vista_FileView(lvFileVw As FileView)

    If PV_Section <> "" And PV_Key <> "" Then
        Save_Config "Configuraciones\" & PV_Section, lvFileVw.Name & ".ViewStyle", lvFileVw.ViewStyle
    End If
    
End Sub

Private Sub Init_Controls_Botones(LV_Path As String)

Dim lv_Path_Iconos  As String
Dim lsFile          As String
Dim i               As Integer

    With UserControl
        lv_Path_Iconos = Get_Config("Configuraciones", "lv_Path_Iconos", "Iconos")
        lv_Path_Iconos = LV_Path & "\" & lv_Path_Iconos & "\"
        
        lsFile = Get_Config("Configuraciones", "Icono SubirNivel", "SubirNivel")
        btn_Nsuperior.Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
        lsFile = Get_Config("Configuraciones", "Icono CarpetaNueva", "CrearCarpeta")
        btn_Nueva.Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
        lsFile = Get_Config("Configuraciones", "Icono Vista", "Vista")
        btn_vistas.Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
        lsFile = Get_Config("Configuraciones", "Icono Select_All", "Select_All")
        .Image_Btn_File_Operation(0).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
        lsFile = Get_Config("Configuraciones", "Icono Deselect", "Deselect")
        .Image_Btn_File_Operation(1).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
        lsFile = Get_Config("Configuraciones", "Icono Move_To", "Move_To")
        .Image_Btn_File_Operation(2).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
        lsFile = Get_Config("Configuraciones", "Icono Remove_From", "Remove_From")
        .Image_Btn_File_Operation(3).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
'        lsFile = Get_Config("Configuraciones", "Icono Add_All", "Add_All")
'        .Image_Btn_File_Operation(0).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
'
'        lsFile = Get_Config("Configuraciones", "Icono Add_All", "Add_All")
'        .Image_Btn_File_Operation(0).Picture = LoadPicture(lv_Path_Iconos & lsFile & ".ico")
    
    End With
    
End Sub

Sub Init_Controls()

Dim LV_Path         As String
Dim lv_Path_Iconos  As String
Dim lsFile          As String

    With UserControl
        '
        Set PC_Boton3D = New clsBoton3D
        
        '
        Set PC_TxtNvaExt = New clsTextBox
        PC_TxtNvaExt.Iniciar .txtNvaExtension, Me
        .txtNvaExtension.Text = ""
        
        ' Filtros de Extension
        Get_Filtros_Extension .cboExtensiones
        
        '
        .PictureBotonesDir.BorderStyle = 0
        
        '
        .ShComboBoxAbPls.FileView = .FileViewAbrirPls
        
        ' Cargar Íconos e Imágenes
        LV_Path = Retroceder_Path(App.Path)
        
        Init_Controls_Botones LV_Path
        
        '
        'lv_Path = Get_Config("Configuraciones", "Ultimo Path", lv_Path)
        LV_Path = Obtener_Ultimo_Path(LV_Path)
        '.ShComboBoxAbPls
        
        '
        .FileViewAbrirPls.CurrentFolder = LV_Path
        Recuperar_Vista_FileVw .FileViewAbrirPls
        
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
    
    If LV_Count Then
        ReDim Preserve lv_File_List(LV_Count - 1)
        LstVw_Add_ListItems lvLstVw, lv_File_List, True
        LVSetAllColWidths lvLstVw, LVSCW_AUTOSIZE
    End If
    
End Sub

Private Sub Recuperar_Vista_FileVw(lvFileVw As FileView)

    If PV_Section <> "" And PV_Key <> "" Then
        lvFileVw.ViewStyle = Get_Config("Configuraciones", lvFileVw.Name & ".ViewStyle", lvFileVw.ViewStyle)
    End If
    
End Sub

Private Sub btn_Nsuperior_Click()

    UserControl.ShComboBoxAbPls.GoUp
    
End Sub

Private Sub btn_Nsuperior_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With UserControl
        PC_Boton3D.Iniciar .btn_Nsuperior, .btn_Nsuperior.Container, .tmrEfecto3D
    End With
    
End Sub

Private Sub btn_Nueva_Click()

    frmNvaCarpeta.Set_Last_Path UserControl.txtPathPulsos.Text
    frmNvaCarpeta.Show vbModal
    UserControl.FileViewAbrirPls.RefreshView
    
End Sub

Private Sub btn_Nueva_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With UserControl
        PC_Boton3D.Iniciar .btn_Nueva, .btn_Nueva.Container, .tmrEfecto3D
    End With

End Sub

Private Sub btn_vistas_Click()

    With UserControl
        .PopupMenu .mnuVer
    End With
    
End Sub

Private Sub btn_vistas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With UserControl
        PC_Boton3D.Iniciar .btn_vistas, .btn_vistas.Container, .tmrEfecto3D
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
    
    Unload Me
    
End Sub

Private Sub FileViewAbrirPls_Click()

    With UserControl
        If .FileViewAbrirPls.SelectedCount = 0 Then
            .cmdOpenDatos.Enabled = False
        Else
            .cmdOpenDatos.Enabled = True
            'Guardar_Ultimo_Path
        End If
    End With
    
End Sub

Private Sub New_Path_Selected()

    Guardar_Ultimo_Path
    Put_Files_On_List UserControl.txtPathPulsos.Text, UserControl.cboExtensiones.Text, UserControl.LstVwFileList
    RaiseEvent PathChanged

End Sub

Private Sub FileViewAbrirPls_OnCurrentFolderChanged(ByVal NewFolder As String)

    New_Path_Selected
    
End Sub

Private Sub FileViewAbrirPls_OnItemClick(ByVal Item As FileViewControl.IListItem, ByVal X As Long, ByVal Y As Long)

    New_Path_Selected
    
End Sub

Private Sub Form_Load()

    'Init_Controls
    
End Sub

Private Sub Image_Btn_File_Operation_Click(Index As Integer)

    'Si es el botón izquierdo...
    
    With UserControl
'        If Button = 1 Then
'            'Efecto de pulsado
'            .Image_Btn_File_Operation(Index).BorderStyle = 1
'            'lblTip.Visible = False
'            'cargar la rutina correspondiente
'            Image_Btn_File_Operation(Index).BorderStyle = 0
'        End If
        Select Case Index
            Case 0
                ' Select All
                LstVw_Set_CheckBox .LstVwFileList, True
            Case 1
                ' Deselect All
                LstVw_Set_CheckBox .LstVwFileList, False
            Case 2
                ' Move To
                LstVw_Move_Chequed_To .LstVwFileList, .LstVwFileSelected
            Case 3
                ' Remove From
                LstVw_Move_Selected_To .LstVwFileSelected, .LstVwFileList
        End Select
    End With

End Sub

Private Sub Image_Btn_File_Operation_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'    With UserControl.Image_Btn_File_Operation(Index)
'        If PV_nBotonFileOp Then
'            If PV_nBotonFileOp <> Index + 1 Then
'                .Container.Cls      'Borrar el efecto anterior
'            Else
'                'Si estamos en el mismo botón, salir
'                Exit Sub
'            End If
'        End If
'        PV_nBotonFileOp = Index + 1
'        'Dibujar el efecto "botón"
'        Efecto3DN E3D_RAISED, .Container, Image_Btn_File_Operation(Index)
'    End With
    PC_Boton3D.Iniciar UserControl.Image_Btn_File_Operation(Index), UserControl.Image_Btn_File_Operation(Index).Container, UserControl.tmrEfecto3D

End Sub

Private Sub LstVwFileList_DblClick()

    LstVw_Check_Selected UserControl.LstVwFileList, True
    RaiseEvent FileSelected
    
End Sub

Private Sub mnuVistaDetalles_Click()

    UserControl.FileViewAbrirPls.ViewStyle = Report
    Guardar_Vista_FileView UserControl.FileViewAbrirPls
    
End Sub

Private Sub mnuVistaIconos_Click()

    UserControl.FileViewAbrirPls.ViewStyle = LargeIcon
    Guardar_Vista_FileView UserControl.FileViewAbrirPls
    
End Sub

Private Sub mnuVistaLista_Click()

    UserControl.FileViewAbrirPls.ViewStyle = List
    Guardar_Vista_FileView UserControl.FileViewAbrirPls
    
End Sub

Private Sub mnuVistaMiniatura_Click()

    UserControl.FileViewAbrirPls.ViewStyle = Thumbnails
    Guardar_Vista_FileView UserControl.FileViewAbrirPls
    
End Sub


Private Sub UserControl_Initialize()

    With UserControl
        
    End With
    
End Sub

Private Sub UserControl_Resize()

Dim lWidth          As Long

    With UserControl
        .FrameSelPath.Height = .Height - 45
        .LstVwFileList.Height = .FrameSelPath.Height - 45 - .LstVwFileList.Top
        .FrameSelPath.Width = .Width - 60
        lWidth = .FrameSelPath.Width - .LstVwFileList.Left - 50
        If lWidth >= 1200 Then
            .LstVwFileList.Width = lWidth
        End If
    End With
        
End Sub
