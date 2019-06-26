VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenPulseDialog 
   Caption         =   "Seleccionar Misión con Pulsos"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   Icon            =   "frmOpenPulseDialog.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   11475
   Begin VB.Frame FrameMisionInfo 
      Caption         =   "Informacion Mision"
      Height          =   5775
      Left            =   6360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin MSComctlLib.ListView LstVwMissionInfo 
         Height          =   5415
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
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
   Begin VB.PictureBox CtlOpenDlgPulsesDir 
      Height          =   5055
      Left            =   120
      ScaleHeight     =   4995
      ScaleWidth      =   6075
      TabIndex        =   2
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "frmOpenPulseDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PV_Last_Path_Clicked            As String
Private PV_Height_Min                   As Long
Private PV_Width_Min                    As Long
'
'

Sub Cerrar_Form()

    Unload Me
    
End Sub

Sub Init_Form()

Dim LV_Path             As String

    On Error Resume Next
    'Load_Form_Params Me
    Me.WindowState = vbMaximized
    If m_Project.IsNewProject = True Then
        LV_Path = Obtener_Ultimo_Path_Form(Me.Name, App.Path)
    Else
        LV_Path = m_Project.GetPulsePath
    End If
    GV_Mdi.Set_Visible_MnuProject SelPathMission
    'Open_Buttom_State False
    
    With Me.CtlOpenDlgPulsesDir
        .LastPath = LV_Path
        .Set_Btn_File_Op_State False
        .Set_CmdOpen_State False
        .Set_File_Settings Find_Folder("Config", App.Path) & "\OpenPrjCfg.txt"
        .App_Path = App.Path
        .Set_LstVw_SelFile_State False
        .Set_NewExtension_State False
        .Set_Open_Dialog_Behavior OpenFolder
        SaveSetting App.Title, .Name, "Set_File_Settings", Find_Folder("Config", App.Path) & "\OpenPrjCfg.txt"
        .Init_Controls
        'WriteLogFile ".Path_Iconos = " & .Path_Iconos
        SaveSetting App.Title, .Name, "Path_Iconos", .Path_Iconos
        SaveSetting App.Title, .Name, "File_Btn_Nsuperior", .File_Btn_Nsuperior
        SaveSetting App.Title, .Name, "HandlerFileLog", .HandlerFileLog
        'WriteLogFile ".File_Btn_Nsuperior =" & .File_Btn_Nsuperior
        'WriteLogFile "Handle File Log Ocx = " & .HandlerFileLog
    End With
    ShowPathInfo LV_Path, Me.LstVwMissionInfo, Me.cmdAbrir
    'On Error GoTo 0
    
End Sub

' evalua lvpath si tiene archivos de pulsos retornando el número de archivos
' en caso de no corresponder a una estructura de misiòn, retorna -1
' en caso de no haber pulsos, retorna 0
' en caso de haber pulsos, retorna cantidad de archivos
Function VerifyMissionPath(ByVal lvPath As String, ByRef lvMissionPath, _
                            ByRef lvPulsePath, ByRef lvTimeIni As String, _
                            ByRef lvTimeEnd As String) As Long

Dim lvPathPls           As String
Dim lvFileCount         As Long

    lvMissionPath = ""
    lvPulsePath = ""
    lvFileCount = -1
    lvTimeIni = Space(260)
    lvTimeEnd = Space(260)
    lvPathPls = lvPath & GC_NORMALES
    If Is_Folder(lvPathPls) = True Then
        '/Pulse_SetMissionPath lvPath
        lvMissionPath = lvPath
        lvPulsePath = lvPathPls
        Pulse_GetMissionInfo lvPath, lvFileCount, lvTimeIni, lvTimeEnd
    Else
        lvPathPls = lvPath
        lvPath = Retroceder_Path(lvPathPls)
        If lvPath & GC_NORMALES = lvPathPls Then
            lvMissionPath = lvPath
            lvPulsePath = lvPathPls
        Pulse_GetMissionInfo lvPath, lvFileCount, lvTimeIni, lvTimeEnd
        Else
            lvTimeIni = ""
            lvTimeEnd = ""
        End If
    End If
    lvTimeIni = Left$(lvTimeIni, 19)
    lvTimeEnd = Left$(lvTimeEnd, 19)
    VerifyMissionPath = lvFileCount
    
End Function

Private Sub cmdAbrir_Click()

Dim lFiles          As Long
Dim lvResult        As Long
Dim lvMissionPath   As String
Dim lvPulsePath     As String
Dim lvTimeIni       As String
Dim lvTimeEnd       As String
Dim lvFilesCount    As Long

    Set_MousePointer vbHourglass
    Me.Enabled = False
    GV_Pulse_Path = Me.CtlOpenDlgPulsesDir.LastPath
    Guardar_Ultimo_Path_Form Me.Name, Me.CtlOpenDlgPulsesDir.LastPath
    lvFilesCount = VerifyMissionPath(GV_Pulse_Path, lvMissionPath, lvPulsePath, _
                        lvTimeIni, lvTimeEnd)
    If lvFilesCount <= 0 Then
        lvFilesCount = VerifyMissionPath(PV_Last_Path_Clicked, lvMissionPath, lvPulsePath, _
                            lvTimeIni, lvTimeEnd)
        If lvFilesCount <= 0 Then
            Exit Sub
        End If
    End If
    'm_Project.SetOutputPath GV_Pulse_Path
    m_Project.SetMissionPath lvMissionPath
    m_Project.SetPulsePath lvPulsePath
    m_Project.SetPulsesTime lvTimeIni, lvTimeEnd
    m_Project.SetFileCount lvFilesCount
    m_Project.SetMissionSelected
    
    If m_Project.IsNewProject = True Then
        GV_clsProjectSelLocation.SetLoad
    End If

'    Set_MousePointer vbHourglass
'
'    GV_WorkSpace = GV_Pulse_Path & "\WrkSpc"
'    GV_Output = GV_Pulse_Path & "\Output"
'    Create_Folder GV_WorkSpace
'    Create_Folder GV_Output
'    lFiles = GetSetting(App.Title, "StructDatos", "FilesPerWorkSpace", 4)
'    Pulse_FilesPerWorkSpace lFiles
'    Pulse_SetWorkSpacePath GV_WorkSpace
'    Pulse_OutputPath GV_Output
'    Pulse_Import_File GV_Pulse_Path
'
'    lvResult = Pulse_Import_File(GV_Pulse_Path)
'    If lvResult = 0 Then
'        Set_MousePointer vbDefault
'        Exit Sub
'    End If
'    If m_Project.IsNewProject = True Then
'        GV_clsConfigSpreadSheet.SetLoad
'    End If
'    GV_MissionName = w(GV_Pulse_Path)
    m_Project.Run_Pulse_Analize
    Set_MousePointer vbDefault
    
    Cerrar_Form
    
End Sub

Private Sub cmdCancelar_Click()

    Cerrar_Form
    If m_Project.IsNewProject = True Then
        m_Project.DiscardProject
        GV_Mdi.Close_Project
    End If
    
End Sub

Sub ShowPathInfo(lvPath As String, LstVw As ListView, LV_Cmd As CommandButton)

Dim lvPathPls           As String

    PV_Last_Path_Clicked = lvPath
    LstVw.ListItems.Clear
    lvPathPls = lvPath & GC_NORMALES
    If Is_Folder(lvPathPls) = True Then
        '/Pulse_SetMissionPath lvPath
        ShowMissionInfo LstVw, lvPath, LV_Cmd
    Else
        lvPathPls = lvPath
        lvPath = Retroceder_Path(lvPathPls)
        If lvPath & GC_NORMALES = lvPathPls Then
            ShowMissionInfo LstVw, lvPath, LV_Cmd
        Else
            LV_Cmd.Enabled = False
        End If
    End If
    
End Sub

Private Sub CtlOpenDlgPulsesDir_PathChanged()

    ShowPathInfo Me.CtlOpenDlgPulsesDir.LastPath, Me.LstVwMissionInfo, _
                    Me.cmdAbrir

End Sub

Sub Set_Form_Caption()

    If m_Project.IsNewProject = True Then
        Me.Caption = "Proyecto Nuevo - Ubicación de Pulsos"
    Else
        Me.Caption = "Seleccionando nueva ubicación de Pulsos"
    End If
    
End Sub

Sub ShowMissionInfo(LstVw As ListView, lvPath As String, LV_Cmd As CommandButton)

Dim lvFileCount         As Long
Dim lvTimeIni           As String
Dim lvTimeEnd           As String
Dim LstItm              As ListItem

    Me.Enabled = False
    Set_MousePointer vbArrowHourglass
    lvTimeIni = Space(260)
    lvTimeEnd = Space(260)
    Pulse_GetMissionInfo lvPath, lvFileCount, lvTimeIni, lvTimeEnd
    With LstVw
        .ListItems.Clear
        Set LstItm = .ListItems.Add(, , "Misión")
        LstItm.ListSubItems.Add , , GetMissionNameFromPath(lvPath & GC_NORMALES)
        Set LstItm = .ListItems.Add(, , "Ubicación")
        LstItm.ListSubItems.Add , , lvPath
        Set LstItm = .ListItems.Add(, , "Archivos")
        LstItm.ListSubItems.Add , , Trim$(Str(lvFileCount))
        Set LstItm = .ListItems.Add(, , "Tiempo Ini")
        LstItm.ListSubItems.Add , , lvTimeIni
        Set LstItm = .ListItems.Add(, , "Tiempo Fin")
        LstItm.ListSubItems.Add , , lvTimeEnd
    End With
    AutoAjusteColumnWidth LstVw
    If lvFileCount Then
        LV_Cmd.Enabled = Not m_Project.GetExportQueued
    Else
        LV_Cmd.Enabled = False
    End If
    Set_MousePointer vbDefault
    Me.Enabled = True
    
End Sub

Private Sub CtlOpenDlgPulsesDir_PathClicked(lvPath As String)

    ShowPathInfo lvPath, Me.LstVwMissionInfo, Me.cmdAbrir
    
End Sub

Private Sub Form_Load()

    PV_Height_Min = 5800
    PV_Width_Min = 11000
    GV_Mdi.Set_Status_MnuProject SelPathMission, False
    Init_Form
    Set_Form_Caption
    Set_MousePointer vbDefault
    'Me.cmdAbrir.Enabled = Not m_Project.GetExportQueued
    
End Sub

Private Sub Form_Resize()

Dim lvGap           As Long

    With Me
        lvGap = 120
        If .ScaleHeight > PV_Height_Min Then
            .FrameMisionInfo.Height = .ScaleHeight - lvGap
            .cmdAbrir.Top = .FrameMisionInfo.Top + .FrameMisionInfo.Height - .cmdAbrir.Height
            .cmdCancelar.Top = .cmdAbrir.Top
            .CtlOpenDlgPulsesDir.Height = .cmdAbrir.Top - 2 * .CtlOpenDlgPulsesDir.Top
            .LstVwMissionInfo.Height = .FrameMisionInfo.Height - _
                                        2 * .LstVwMissionInfo.Top
        End If
        If .ScaleWidth > PV_Width_Min Then
            .FrameMisionInfo.Width = .ScaleWidth - .FrameMisionInfo.Left - lvGap
            .LstVwMissionInfo.Width = .FrameMisionInfo.Width - _
                                        2 * .LstVwMissionInfo.Left
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GV_clsOpenPulseDialog.ClearLoaded
    'GV_Mdi.Restore_Visible_MnuProject SelPathMission
    Save_Form_Params Me
    GV_Mdi.ProjectMnuUpdate
    
End Sub
