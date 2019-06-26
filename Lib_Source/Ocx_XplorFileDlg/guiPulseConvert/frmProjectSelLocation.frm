VERSION 5.00
Begin VB.Form frmProjectSelLocation 
   Caption         =   "Seleccionar Ubicacion del Proyecto"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   10485
   Begin VB.Frame FrameProjectSelLocation 
      Caption         =   "Información del Proyecto"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtPulseFileCount 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         TabIndex        =   34
         Text            =   "txtPulseFileCount"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.PictureBox pictureNewData 
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   9375
         TabIndex        =   16
         Top             =   3000
         Width           =   9375
         Begin VB.CheckBox chkCreateFolderForProject 
            Caption         =   "Crear Carpeta con el Nombre del Proyecto"
            Height          =   255
            Left            =   3720
            TabIndex        =   35
            Top             =   0
            Width           =   3375
         End
         Begin VB.CheckBox chkNameAssociatedToMission 
            Caption         =   "Asociar Nombre del Proyecto con la Misión"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   3855
         End
         Begin VB.TextBox txtProjectName 
            Height          =   285
            Left            =   2280
            TabIndex        =   27
            Text            =   "txtProjectName"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtPjtPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   26
            Text            =   "txtPjtPath"
            Top             =   960
            Width           =   8895
         End
         Begin VB.CommandButton cmdSelPath 
            Caption         =   "..."
            Height          =   195
            Index           =   0
            Left            =   9000
            TabIndex        =   25
            Top             =   1020
            Width           =   375
         End
         Begin VB.TextBox txtPjtOutputPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   24
            Text            =   "txtPjtOutputPath"
            Top             =   1800
            Width           =   8895
         End
         Begin VB.TextBox txtPjtWorkSpacePath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   23
            Text            =   "txtPjtWorkSpacePath"
            Top             =   2520
            Width           =   8895
         End
         Begin VB.CommandButton cmdSelPath 
            Caption         =   "..."
            Height          =   195
            Index           =   1
            Left            =   9000
            TabIndex        =   22
            Top             =   1920
            Width           =   375
         End
         Begin VB.CommandButton cmdSelPath 
            Caption         =   "..."
            Height          =   195
            Index           =   2
            Left            =   9000
            TabIndex        =   21
            Top             =   2640
            Width           =   375
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   0
            TabIndex        =   20
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CommandButton cmdAccept 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   8160
            TabIndex        =   19
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CheckBox chkLinkOutputPath 
            Caption         =   "Carpeta de Salida como Sub Carpeta del Proyecto"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2760
            TabIndex        =   18
            Top             =   1440
            Width           =   4575
         End
         Begin VB.CheckBox chkLinkWorkSpacePath 
            Caption         =   "Carpeta de Archivos Intermedios como Sub Carpeta del Proyecto"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2760
            TabIndex        =   17
            Top             =   2160
            Width           =   5295
         End
         Begin VB.Label lblProjectSelLocation 
            Caption         =   "Ubicación Nuevo Proyecto:"
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lbl 
            Caption         =   "Nombre del Proyecto:"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lbl 
            Caption         =   "Carpeta Archivos de Salida:"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   30
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lbl 
            Caption         =   "Carpeta Archivos Intermedios:"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   29
            Top             =   2160
            Width           =   2175
         End
      End
      Begin VB.PictureBox pictureMissionName 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1800
         ScaleHeight     =   375
         ScaleWidth      =   4215
         TabIndex        =   14
         Top             =   360
         Width           =   4215
         Begin VB.TextBox txtMissionName 
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   15
            Text            =   "txtMissionName"
            Top             =   0
            Width           =   4215
         End
      End
      Begin VB.PictureBox picturePulsePath 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9375
         TabIndex        =   12
         Top             =   2520
         Width           =   9375
         Begin VB.TextBox txtPulsePath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   13
            Text            =   "txtPulsePath"
            Top             =   0
            Width           =   9375
         End
      End
      Begin VB.PictureBox pictureMissionPath 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   9375
         TabIndex        =   10
         Top             =   1800
         Width           =   9375
         Begin VB.TextBox txtMissionPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   11
            Text            =   "txtMissionPath"
            Top             =   0
            Width           =   9375
         End
      End
      Begin VB.TextBox txtPulseEndTime 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         TabIndex        =   9
         Text            =   "txtPulseEndTime"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtPulseIniTime 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Text            =   "txtPulseIniTime"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtMissionDate 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         TabIndex        =   7
         Text            =   "txtMissionDate"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Cantidad de Archivos de Pulso:"
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
         Index           =   8
         Left            =   4560
         TabIndex        =   33
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lbl 
         Caption         =   "Tiempo Término de Pulsos:"
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
         Index           =   7
         Left            =   4800
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lbl 
         Caption         =   "Tiempo Inicio de Pulsos:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lbl 
         Caption         =   "Fecha Misión:"
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
         Index           =   5
         Left            =   6120
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl 
         Caption         =   "Ubicación Pulsos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblMissionPath 
         Caption         =   "Ubicación Misión:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lbl 
         Caption         =   "Nombre Misión:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmProjectSelLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmProjectSelLocation
' Author    : Leo Herrera
' Date      : 12/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Dim mPjtName            As clsSelText

Sub AsociarMissionWithProjectName(lvState As Boolean)

    If lvState = True Then
        Me.txtProjectName = Me.txtMissionName
        Me.txtProjectName.Enabled = False
    Else
        Me.txtProjectName.Enabled = True
    End If
    
End Sub

Sub Init_Classes()

    Set mPjtName = New clsSelText
    mPjtName.SetControl Me.txtProjectName
    
End Sub

Sub Init_Info_Mission()

    With Me
        .Init_Classes
        .txtMissionDate = m_Project.GetMissionDate
        .txtMissionName = m_Project.GetMissionName
        .pictureMissionName.ToolTipText = m_Project.GetMissionName
        .txtMissionPath = m_Project.GetMissionPath
        .pictureMissionPath.ToolTipText = m_Project.GetMissionPath
        .txtPulseEndTime = m_Project.GetPulseEndTime
        .txtPulseIniTime = m_Project.GetPulseIniTime
        .txtPulsePath = m_Project.GetPulsePath
        .picturePulsePath.ToolTipText = m_Project.GetPulsePath
        .txtPulseFileCount = m_Project.GetFilesCount
    End With
    
End Sub

Sub Init_Form_Old_Project()

    On Error Resume Next
    With Me
        .Init_Info_Mission
        .chkNameAssociatedToMission.Value = m_Project.GetAsociaMissionName
        .chkLinkOutputPath.Value = m_Project.GetOutputSubFolder
        .chkLinkWorkSpacePath.Value = m_Project.GetWrkSpcSubFolder
        .chkCreateFolderForProject = m_Project.GetCreateFolderProject
        GV_Mdi.Set_Visible_MnuProject SelPathProject
        .txtPjtOutputPath = m_Project.GetOutputPath
        .txtPjtPath = m_Project.GetProjectPath
        .txtPjtWorkSpacePath = m_Project.GetWorkSpacePath
        .txtProjectName = m_Project.GetName
    End With
    'On Error GoTo 0
    
End Sub

Sub Init_Form()

    With Me
        .Init_Info_Mission
        GV_Mdi.Set_Visible_MnuProject SelPathProject
        GetSettingCheckBox .chkNameAssociatedToMission
        .txtPjtOutputPath = ""
        .txtPjtPath = ""
        .txtPjtWorkSpacePath = ""
        .txtProjectName = ""
        GetSettingCheckBox .chkLinkOutputPath
        GetSettingCheckBox .chkLinkWorkSpacePath
        GetSettingCheckBox .chkNameAssociatedToMission
        GetSettingCheckBox .chkCreateFolderForProject
        If .chkNameAssociatedToMission.Value Then
            AsociarMissionWithProjectName True
        Else
            AsociarMissionWithProjectName False
        End If
    End With
    
End Sub

Private Sub chkCreateFolderForProject_Click()

    SaveSettingCheckBox Me.chkCreateFolderForProject
    
End Sub

Private Sub chkLinkOutputPath_Click()

    With Me
        SaveSettingCheckBox .chkLinkOutputPath
        If .chkLinkOutputPath.Value Then
            .cmdSelPath(1).Enabled = False
            .AutoCompletarPath
        Else
            .cmdSelPath(1).Enabled = True
        End If
    End With
    
End Sub

Private Sub chkLinkWorkSpacePath_Click()

    With Me
        SaveSettingCheckBox .chkLinkWorkSpacePath
        If .chkLinkWorkSpacePath.Value Then
            .cmdSelPath(2).Enabled = False
            .AutoCompletarPath
        Else
            .cmdSelPath(2).Enabled = True
        End If
    End With
End Sub

Private Sub chkNameAssociatedToMission_Click()

    SaveSettingCheckBox Me.chkNameAssociatedToMission
    If Me.chkNameAssociatedToMission.Value Then
        AsociarMissionWithProjectName True
    Else
        AsociarMissionWithProjectName False
    End If
    
End Sub

Private Sub cmdAccept_Click()

'Dim lvForm          As frmDialogAnalizingData
Dim lvTickIni, lvEnlased            As Long
Dim lvIterations        As Long
Dim i, j              As Long
Dim LV_Form_Prepro      As frmPreprocessing
            
    With Me
        If .txtProjectName = "" Then
            Exit Sub
        End If
        If .txtPjtPath = "" Then
            Exit Sub
        End If
        If .txtPjtOutputPath = "" Then
            Exit Sub
        End If
        If .txtPjtWorkSpacePath = "" Then
            Exit Sub
        End If
        m_Project.SetProjectPath .txtPjtPath
        m_Project.SetOutputPath .txtPjtOutputPath
        m_Project.SetWorkSpacePath .txtPjtWorkSpacePath
        m_Project.SetName .txtProjectName
        m_Project.CreateFolders
        If m_Project.SaveProject = True Then
            modProjectFunctions.AddProjectToList
        Else
            Exit Sub
        End If
        m_Project.SetProjectFolderSelected
        lvIterations = GetSetting(App.Title, "Debugging", "Run_Pulse_Analize_Iterations", 0)
        SaveSetting App.Title, "Debugging", "Run_Pulse_Analize_Iterations", lvIterations
        
        Set LV_Form_Prepro = New frmPreprocessing
        GV_Mdi.Enabled = False
        LV_Form_Prepro.Show vbModal
        'LV_Form_Prepro.Refresh
        'doevents
        If lvIterations = 0 Then
            'm_Project.Run_Pulse_Analize
            m_Project.Set_Parameters_After_Pulse_Analize
        Else
            'modLog.OpenLogFile
            For j = 1 To 8
            m_Project.SetFilesPerWorkSpace 128
            For i = 1 To lvIterations
                lvTickIni = GetTickCount
                m_Project.Run_Pulse_Analize
                lvEnlased = GetTickCount - lvTickIni
                
                'modLog.WriteLogFile "FilesPerWorkSpace;" & m_Project.GetFilesPerWorkSpace _
                                    & ";Tpo Ini;" & lvTickIni _
                                    & ";Tpo Enlasado;" & lvEnlased / 1000
            Next
            m_Project.SetFilesPerWorkSpace m_Project.GetFilesPerWorkSpace / 2
            Next
            'modLog.CloseLogFile
            m_Project.SetFilesPerWorkSpace GetSettingFilesPerWorkSpace
        End If
        
'        Load lvForm
'        lvForm.Show vbModal
        
        If m_Project.SaveProject(True) = False Then
            Exit Sub
        End If
        If m_Project.GetPulsesAnalized = True Then
            GV_clsTemplateConfigSpreadSheet.SetLoad
        End If
        GV_Mdi.Enabled = True
        Unload LV_Form_Prepro
    End With
    
    Unload Me
    
End Sub

Function CalcPjtPath(lvPath As String, lvChk As CheckBox, LvPjtName As TextBox) As String

    If lvChk.Value Then
        CalcPjtPath = lvPath & "\" & LvPjtName
    Else
        CalcPjtPath = lvPath
    End If
    
End Function

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdSelPath_Click(Index As Integer)

Dim sDir        As String
Dim lFlags      As Long
Dim lPath       As String
Dim sFile       As String

    lFlags = BIF_RETURNONLYFSDIRS
    lPath = GetSetting(App.Title, Me.Name, "Sel Dir " & Trim$(Str(Index)), "")
    
    sDir = BrowseForFolder(Me.hwnd, "Seleccionar Directorio", lPath, lFlags)
    'sDir = CommonDialogForFolder(Me.CommonDialog, "Seleccionar Directorio", lPath)
    If Err = 0 And sDir <> "" Then
        SaveSetting App.Title, Me.Name, "Sel Dir " & Trim$(Str(Index)), sDir
        Select Case Index
            Case Is = 0
                Me.txtPjtPath = CalcPjtPath(sDir, _
                                            Me.chkCreateFolderForProject, _
                                            Me.txtProjectName)
                Me.chkLinkOutputPath.Enabled = True
                Me.chkLinkWorkSpacePath.Enabled = True
                Me.AutoCompletarPath
            Case Is = 1
                Me.txtPjtOutputPath = sDir
            Case Is = 2
                Me.txtPjtWorkSpacePath = sDir
        End Select
    Else
        'MsgBox "Se ha cancelado la operación, el error devuelto es:" & vbCrLf & _
               "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
        Err = 0
    End If

End Sub

Sub AutoCompletarPath()

    With Me
        If .txtPjtPath.Text <> "" Then
            If .chkLinkWorkSpacePath.Value Then
                .txtPjtWorkSpacePath = .txtPjtPath & "\WrkSpace"
            End If
            If .chkLinkOutputPath.Value Then
                .txtPjtOutputPath = .txtPjtPath & "\Output"
            End If
        End If
    End With

End Sub

Private Sub Form_Load()

    GV_Mdi.Set_Status_MnuProject SelPathProject, False
    With Me
        .WindowState = vbMaximized
        If m_Project.IsNewProject = True Then
            .Init_Form
        Else
            .Init_Form_Old_Project
        End If
    End With
    'Me.cmdAccept.Enabled = Not m_Project.GetExportQueued
    Set_MousePointer vbDefault

End Sub

Private Sub Form_Resize()

    With Me
        If .ScaleHeight > .cmdAccept.Top + .cmdAccept.Height Then
            .FrameProjectSelLocation.Height = .ScaleHeight - _
                                            2 * .FrameProjectSelLocation.Top
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GV_clsProjectSelLocation.ClearLoaded
    GV_Mdi.ProjectMnuUpdate
    
End Sub

Sub VerificarBotonAceptar(LV_TxtBx1 As TextBox, _
                        LV_TxtBx2 As TextBox, _
                        LV_TxtBx3 As TextBox, _
                        LV_Cmd As CommandButton)

    If LV_TxtBx1.Text = "" Or LV_TxtBx2.Text = "" Or LV_TxtBx3.Text = "" Then
        LV_Cmd.Enabled = False
    Else
        LV_Cmd.Enabled = Not m_Project.GetExportQueued
    End If
    
End Sub

Private Sub txtPjtOutputPath_Change()

    With Me
        .VerificarBotonAceptar .txtPjtOutputPath, .txtPjtPath, .txtPjtWorkSpacePath, .cmdAccept
    End With

End Sub

Private Sub txtPjtPath_Change()

    With Me
        .AutoCompletarPath
        .VerificarBotonAceptar .txtPjtOutputPath, .txtPjtPath, .txtPjtWorkSpacePath, .cmdAccept
    End With
    
End Sub

Private Sub txtPjtWorkSpacePath_Change()

    With Me
        .AutoCompletarPath
        .VerificarBotonAceptar .txtPjtOutputPath, .txtPjtPath, .txtPjtWorkSpacePath, .cmdAccept
    End With
    
End Sub

Private Sub txtProjectName_Change()

    With Me
        If .txtPjtPath.Enabled = True Then
            .AutoCompletarPath
        End If
        .VerificarBotonAceptar .txtPjtOutputPath, .txtPjtPath, .txtPjtWorkSpacePath, .cmdAccept
    End With
    
End Sub
