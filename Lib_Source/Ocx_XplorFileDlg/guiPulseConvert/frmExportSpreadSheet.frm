VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportSpreadSheet 
   Caption         =   "Exportanto Pulsos a Excel"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   Icon            =   "frmExportSpreadSheet.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   13080
   Begin VB.PictureBox pictureDeprecated 
      Height          =   4695
      Left            =   7320
      ScaleHeight     =   4635
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txtDbg 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Text            =   "frmExportSpreadSheet.frx":000C
         Top             =   3840
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox txtSckData 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "frmExportSpreadSheet.frx":0013
         Top             =   2160
         Width           =   6615
      End
      Begin VB.CommandButton cmdAccept 
         Caption         =   "&Comenzar"
         Height          =   375
         Left            =   1320
         TabIndex        =   19
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox chkRunInBackGround 
         Caption         =   "Correr en Background"
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   3360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FrameSpreadSheetStatus 
         Caption         =   "Estado de Avance"
         Height          =   1935
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   5895
         Begin VB.CommandButton cmdPause 
            Caption         =   "Pausar"
            Height          =   255
            Left            =   4800
            TabIndex        =   7
            Top             =   1560
            Width           =   855
         End
         Begin VB.CommandButton cmdCancelProcess 
            Caption         =   "Cancelar"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   855
         End
         Begin MSComctlLib.ProgressBar ProgressBarAvance 
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   8
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBarAvance 
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   9
            Top             =   720
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBarAvance 
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   10
            Top             =   1080
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblPulsos 
            Caption         =   "Pulsos:"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblArchivos 
            Caption         =   "Hojas:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblProceso 
            Caption         =   "Archivos:"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbl1 
            Caption         =   "Tiempo Enlasado:"
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
            Left            =   1320
            TabIndex        =   12
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblEnlasedTime 
            Caption         =   "lblEnlasedTpo"
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
            Left            =   3240
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
         End
      End
   End
   Begin VB.Frame FrameActualProject 
      Caption         =   "Proyecto en Edición:"
      Height          =   2895
      Left            =   120
      TabIndex        =   22
      Top             =   3840
      Width           =   9615
      Begin MSComctlLib.ListView LstVwActualProject 
         Height          =   2055
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
      Begin VB.PictureBox pictureControlBar 
         BorderStyle     =   0  'None
         Height          =   460
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   4215
         TabIndex        =   23
         Top             =   240
         Width           =   4215
         Begin MSComctlLib.Toolbar toolbarControles 
            Height          =   390
            Left            =   3360
            TabIndex        =   24
            Top             =   0
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            ImageList       =   "ImageListControles"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageListControles 
            Left            =   3720
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExportSpreadSheet.frx":001E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExportSpreadSheet.frx":0A30
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExportSpreadSheet.frx":1442
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExportSpreadSheet.frx":19DC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExportSpreadSheet.frx":1F76
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExportSpreadSheet.frx":2510
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmExportSpreadSheet.frx":266A
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label lblStart 
            Caption         =   "Agregar a la Lista de Exportación"
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
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   3615
         End
      End
   End
   Begin VB.PictureBox pictureOutputPath 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.Timer TimerCreating 
         Enabled         =   0   'False
         Left            =   6480
         Top             =   0
      End
      Begin VB.Timer TimerSaving 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6960
         Top             =   0
      End
      Begin VB.TextBox txtProcessingProject 
         Height          =   285
         Left            =   1440
         TabIndex        =   21
         Text            =   "txtProcessingProject"
         Top             =   120
         Width           =   4815
      End
      Begin VB.Frame FrameResults 
         Caption         =   "Generar Xls"
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   12855
         Begin MSComctlLib.ListView LstVwExport 
            Height          =   3015
            Left            =   3240
            TabIndex        =   3
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   5318
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
               Text            =   "Parámetro"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Valor"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.TreeView TrVwExport 
            Height          =   3015
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   5318
            _Version        =   393217
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
      End
      Begin VB.Label lbl 
         Caption         =   "Procesando:"
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
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmExportSpreadSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmExportSpreadSheet
' Author    : Leo Herrera
' Date      : 27/10/2012
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

Dim PV_New_Msg              As Boolean
Dim PV_Index_Old(2)         As Long
Dim PV_FileSavingSent       As Boolean
Dim PV_Height_Min           As Long
Dim PV_Width_Min            As Long

Private PV_RunCvtXls        As Boolean

'---------------------------------------------------------------------------------------
'
' me.TrVwExport
' - Project Name
'   - File 1 de Na
'       - Sheet 1 de Ns
'         - Pulsos 1 de Np
'
' Me.LstVwExport
' ->Project Name
' Project Name
' Project Folder
' Output Folder
' Total Files
' Files Exported
' Actual Job

Private PV_PulseCountArray()    As Long
Private PV_PulseAcumArray()     As Double
Private PV_PulseTotal           As Double
Private PV_FileName             As String
Private PV_IndexSpread          As Long
Private PV_State                As Integer

Private PV_Ini_Time             As Date

Private PV_TrVwIndex_BgPrjActive    As Long
Private PV_TrVwIndex_Selected  As Long

Private PV_hFile                As Integer

Sub ShowData(lsData As String, Optional lsFunction As String = "")

Dim lsName          As String
Dim lsStr           As String

   ''On Error GoTo ShowData_Error

    If GetSettingBooleanParameter(GC_ENABLE_EXPORTFILE_LOG, False) = True Then
        If PV_hFile = 0 Then
            If GV_PrevInstance = False Then
                lsName = "Dbg_Main_" & Format(Now(), "hh_mm_ss") & ".log"
            Else
                lsName = "Dbg_Scnd_" & Format(Now(), "hh_mm_ss") & ".log"
            End If
            lsName = Retroceder_Path(App.Path) & "\Exe\" & lsName
            PV_hFile = FreeFile
            Open lsName For Append As PV_hFile
        End If
        If lsFunction <> "" Then
            lsStr = Format(Now(), "hh:mm:ss ") & "| Fn " & lsFunction & "->" & lsData
        Else
            lsStr = Format(Now(), "hh:mm:ss -> ") & lsData
        End If
        Print #PV_hFile, lsStr
    End If
    With Me.txtDbg
        If Len(.Text) > 20000 Then
            .Text = Mid$(.Text, 10000)
        End If
        .SelStart = Len(.Text)
        .SelText = lsData & vbCrLf
        .Refresh
    End With

   'On Error GoTo 0
   Exit Sub

ShowData_Error:

    PV_hFile = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowData of Formulario frmExportSpreadSheet"
    
End Sub

Function VerifyNewInformation(IndexSpread As Long, _
                              IndexSheet As Long, _
                              IndexPulse As Long) As Boolean

    VerifyNewInformation = False
    If PV_Index_Old(2) <> IndexPulse Then
        VerifyNewInformation = True
        PV_Index_Old(2) = IndexPulse
    End If
    If PV_Index_Old(1) <> IndexSheet Then
        VerifyNewInformation = True
        PV_Index_Old(1) = IndexSheet
    End If
    If PV_Index_Old(0) <> IndexSpread Then
        VerifyNewInformation = True
        PV_Index_Old(0) = IndexSpread
    End If
    
End Function

Sub ShowSckData(ByVal lsStr As String)

    With Me.txtSckData
        If Len(.Text) > 60000 Then
            .Text = ""
        End If
        lsStr = lsStr & vbCrLf
        .SelStart = Len(.Text)
        .SelText = lsStr
    End With
    
End Sub

Sub RefreshBgProject()

Dim i           As Long

    ShowData "RefreshBgProject"

    If PV_TrVwIndex_Selected <> PV_TrVwIndex_BgPrjActive Then
        Exit Sub
    End If
    i = FindIndexBgProject(PV_TrVwIndex_Selected)
    If BackGroundProjectList.ProjectList(i).ProjectState = MSG_START_PROJECT Then
        Me.ShowBgProject
    Else
        modBgproject.RefreshBgProject Me.LstVwExport, i
    End If
    
End Sub

Sub EndingCreation()

   '''On Error GoTo EndingCreation_Error

    Pulse_Close_Log
    m_Project.SetSheetGenerated
    m_Project.ClearSheetGenerating
    m_Project.ClearExportQueued
    Me.Caption = "Exportanto Pulsos a Excel - Exportación concluida con éxito"
    m_Project.SaveProject
    Me.SendXlsEndProject
    If GV_PrevInstance = False Then
        If MsgBox("¿Desea cerrar el proyecto?", vbYesNo, "Exportación concluido con éxito") = vbYes Then
            'm_Project.CloseProject
            GV_Mdi.Close_Project
            Unload Me
        End If
    Else
        GV_Mdi.Close_Project
        Unload Me
    End If

   'On Error GoTo 0
   Exit Sub

EndingCreation_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EndingCreation of Formulario frmExportSpreadSheet"
    
End Sub



Sub SendMessage(lsMsg As String)

    GV_Mdi.SendMessage "START_MSG," & lsMsg & ",END_MSG"
    
End Sub


Sub SendXlsFileName(lvFileName As String, _
                    IndexSpread As Long, _
                    lvSheetsCount As Long)

Dim lvMsg           As String
    If GV_Send_Progress = True Then
        With Me
            PV_FileSavingSent = False
            lvMsg = "FILE_START," & lvFileName & "," & IndexSpread & _
                    "," & lvSheetsCount
            Me.SendMessage lvMsg
        End With
    End If
    
End Sub

Sub SendXlsSaving(IndexSpread As Long)

Dim lvMsg           As String

    If PV_FileSavingSent = False Then
        PV_FileSavingSent = True
        lvMsg = GC_MSG_SAVING_FILE & "," & IndexSpread
        Me.SendMessage lvMsg
    End If
    
End Sub

Sub SendXlsFileReady(IndexSpread As Long)

Dim lsMsg           As String

    If GV_Send_Progress = True Then
        With Me
            lsMsg = "XLS_FILE_READY," & IndexSpread
            Me.SendMessage lsMsg
        End With
    End If
    
End Sub

Sub SendXlsStartProject()

Dim lvMsg           As String

    lvMsg = "START_PROJECT," & Pulse_GetSpreadFileCount
    Me.SendMessage lvMsg
    
End Sub

Sub SendXlsEndProject()

Dim lvMsg           As String

    lvMsg = "END_PROJECT"
    Me.SendMessage lvMsg
    
End Sub

Sub SendXlsProgress(IndexSpread As Long, IndexSheet As Long, IndexPulse As Long)

Dim lvData          As String
Dim i               As Integer

    If VerifyNewInformation(IndexSpread, IndexSheet, IndexPulse) = False Then
        Exit Sub
    End If
    If GV_Send_Progress = True Then
        With Me
            lvData = "STATUS," & IndexSpread
            lvData = lvData & "," & IndexSheet
            lvData = lvData & "," & IndexPulse
            Me.SendMessage lvData
'            lvData = "START_MSG,STATUS,"
'            For i = 0 To .ProgressBarAvance.ubound
'                lvData = lvData & .ProgressBarAvance(i).Value & ","
'                lvData = lvData & .ProgressBarAvance(i).Max & ","
'            Next
'            lvData = lvData & "END_MSG"
'            GV_Mdi.SendMessage lvData
        End With
    End If
    
End Sub


Private Sub chkRunInBackGround_Click()

    If GV_PrevInstance = False Then
        SaveSettingCheckBox Me.chkRunInBackGround
    End If
    
End Sub

Private Sub cmdAccept_Click()

Dim lvPath          As String
Dim lvExe           As String

    PV_FileName = m_Project.GetOutputPath & "\" & m_Project.GetName & "_.xls"
    m_Project.SendFieldFormat
    Me.cmdAccept.Enabled = False
    'Pulse_Export_File PV_FileName
    If Me.chkRunInBackGround.Value Then
        'GV_Mdi
        With BackGroundProjectList
            ReDim Preserve .ProjectList(.Count)
            With .ProjectList(.Count)
                .FileName = PV_FileName
                .GenerationDone = False
                .IndexPulse = 0
                .IndexSheet = 0
                .IndexSpread = 0
                .OutFilesCount = 0
                .OutputPath = mProject.OutputPath
                .ProjectName = mProject.ProjectName
                .ProjectPath = mProject.ProjectPath
                .ProjectStarted = False
                .TimeIni = Now()
                .WorkSpacePath = mProject.WorkSpacePath
            End With
            If .Count Then
                AddBgProject Me.TrVwExport, _
                            .ProjectList(.Count), _
                            .ProjectList(.Count - 1).IndexTrVw
            Else
                AddBgProject Me.TrVwExport, .ProjectList(.Count), 0
                PV_TrVwIndex_Selected = 1
                Me.TrVwExport.Nodes(1).Selected = True
                Me.TrVwExportItemSelected
            End If
            'm_Project.ClearSheetGenerated
            'm_Project.SetSheetGenerating
            m_Project.SetExportQueued
            .Count = .Count + 1
            'Me.RefreshList
        End With
    Else
        PV_State = 0
        PV_IndexSpread = 0
        
        Me.TimerCreating.Interval = 50
        Me.TimerCreating.Enabled = True
        Me.FrameSpreadSheetStatus.Visible = True
        Me.cmdAccept.Enabled = False
        Me.cmdCancel.Enabled = False
        
        'InitProgressBarStatus
        With Me.ProgressBarAvance(0)
            .Min = 0
            .Value = 0
            .Max = Pulse_GetSpreadFileCount
        End With
        Me.SendXlsStartProject
        PV_Ini_Time = Now()
        GV_Mdi.RefreshTime
        'modLog.OpenLogFile
    End If
    
End Sub

Sub SetProgressBarStatus(IndexFile As Integer, Pls As Long)

Dim dResult         As Double

    With Me
        If PV_PulseTotal Then
            dResult = PV_PulseAcumArray(IndexFile) + Pls + 1
            dResult = 100 * dResult / PV_PulseTotal
            .ProgressBarAvance(0).Value = dResult
        End If
        .ProgressBarAvance(1).Value = IndexFile + 1
        If PV_PulseCountArray(IndexFile) Then
            .ProgressBarAvance(2).Value = 100 * ((Pls + 1) / PV_PulseCountArray(IndexFile))
        End If
    End With
    
End Sub

Sub InitProgressBarStatus()

Dim dValuePls       As Double
Dim lFileCount      As Integer
Dim i               As Integer

    lFileCount = Pulse_GetSpreadFileCount
    
    ReDim PV_PulseCountArray(lFileCount - 1)
    ReDim PV_PulseAcumArray(lFileCount - 1)
    
    For i = 0 To lFileCount - 1
        PV_PulseCountArray(i) = Pulse_Count(i)
        If i Then
            PV_PulseAcumArray(i) = PV_PulseTotal
        End If
        PV_PulseTotal = PV_PulseTotal + PV_PulseCountArray(i)
    Next
    
    With Me
        For i = 0 To .ProgressBarAvance.UBound
            .ProgressBarAvance(i).Min = 0
            .ProgressBarAvance(i).Value = 0
        Next
        .ProgressBarAvance(0).Max = 100
        .ProgressBarAvance(1).Max = lFileCount
        .ProgressBarAvance(2).Max = 100
    End With

End Sub


Public Sub ParseMsg(lsMsg() As String, IndexMsg As Message_Header_Const)

Dim i               As Integer
Dim IndexSpread     As Long
Dim SheetCount      As Long

    If Verify_Length_Msg(lsMsg, IndexMsg) = False Then
        Exit Sub
    End If
    If BackGroundProjectList.ListIndex = -1 Then
        Exit Sub
    End If
    With BackGroundProjectList.ProjectList(BackGroundProjectList.ListIndex)
        Select Case IndexMsg
            Case Is = Message_Header_Const.MSG_ERROR
            
            Case Is = Message_Header_Const.MSG_START_PROJECT
                If IsNumeric(lsMsg(1)) = True Then
                    Me.txtProcessingProject.Text = .ProjectName
                    .Processing = True
                    .ProjectState = Message_Header_Const.MSG_START_PROJECT
                    .OutFilesCount = lsMsg(1)
                    ReDim .OutFiles(.OutFilesCount - 1)
                    ReDim .OutFilesSheetCount(.OutFilesCount)
                    PV_TrVwIndex_BgPrjActive = .IndexTrVw
                End If

            Case Is = Message_Header_Const.MSG_FILE_START
                If IsNumeric(lsMsg(2)) = True And IsNumeric(lsMsg(3)) = True Then
                    .ProjectState = MSG_FILE_START
                    IndexSpread = Val(lsMsg(2))
                    SheetCount = Val(lsMsg(3))
                    .IndexSpread = IndexSpread
                    .IndexSheet = -1
                    .IndexPulse = -1
                    If .OutFilesCount <= IndexSpread Then
                        .OutFilesCount = IndexSpread + 1
                        ReDim Preserve .OutFiles(IndexSpread)
                        ReDim Preserve .OutFilesSheetCount(IndexSpread)
                    End If
                    .OutFiles(IndexSpread) = GetFileName(lsMsg(1))
                    .OutFilesSheetCount(IndexSpread) = SheetCount
                End If
            Case Is = Message_Header_Const.MSG_STATUS
                .Status = True
                .ProjectState = MSG_STATUS
                .IndexSpread = lsMsg(1)
                .IndexSheet = lsMsg(2)
                .IndexPulse = lsMsg(3)
                UpDateBackgroundProcess BackGroundProjectList.ListIndex
                '.IndexSpreadMax = lsMsg(1)
                '.IndexSheetMax = lsMsg(3)
                '.IndexPulseMax = lsMsg(5)
            Case Is = Message_Header_Const.MSG_SAVING_FILE
                .ProjectState = MSG_SAVING_FILE
            Case Is = Message_Header_Const.MSG_XLS_FILE_READY
                .ProjectState = MSG_XLS_FILE_READY
            Case Is = Message_Header_Const.MSG_END_PROJECT
                .ProjectState = MSG_END_PROJECT
                .GenerationDone = True
'                With BackGroundProjectList
'                    With .ProjectList(.ListIndex)
'                        .GenerationDone = True
'                    End With
'                    .ListIndex = -1
'                End With
        End Select
        Me.RefreshBgProject
    End With
    
End Sub

Sub ShowBgProject()

    ShowData "ShowBgProject"
    If PV_TrVwIndex_Selected = PV_TrVwIndex_BgPrjActive Then
        modBgproject.ShowBgProject Me.LstVwExport, PV_TrVwIndex_BgPrjActive
    End If

End Sub

Sub UpDateBackgroundProcess(Index As Long)

End Sub

Private Sub cmdCancel_Click()

    GV_clsExportSpreadSheet.ClearLoaded
    Unload Me
    
End Sub

Sub UpdateActualProject()

    With Me
        ShowActualProject .LstVwActualProject
        .toolbarControles.Buttons(1).Enabled = Not m_Project.GetExportQueued
    End With
    
End Sub

Private Sub Form_Deactivate()

    GV_clsExportSpreadSheet.ClearLoaded
    
End Sub

Private Sub Form_GotFocus()

    Me.cmdAccept.Enabled = True
    UpdateActualProject
    If m_Project.GetSheetConfigured = True Then
        If m_Project.GetExportQueued = False And _
            m_Project.GetSheetGenerating = False Then
            Me.toolbarControles.Buttons(1).Enabled = True
        End If
    End If
    
End Sub

Private Sub Form_Load()

Dim LV_Path         As String


    GV_Mdi.Set_Status_MnuProject GenOutput, False
    GV_Mdi.Restore_Visible_All_Mnu_For_Open
    With Me
        '.pictureDeprecated.Visible = True
        .txtDbg = ""
        PV_Height_Min = 6800
        PV_Width_Min = 7000
        .txtSckData.Text = ""
        .txtProcessingProject.Text = ""
        .WindowState = vbMaximized
        LV_Path = m_Project.GetOutputPath
        
        GetSettingCheckBox .chkRunInBackGround
        .lblEnlasedTime.Caption = ""
        If .chkRunInBackGround.Value = 0 Then
            .chkRunInBackGround.Value = 1
        End If
        .TrVwExport.Nodes.Clear
        .LstVwExport.ListItems.Clear
        PV_TrVwIndex_BgPrjActive = 0
        PV_TrVwIndex_Selected = 0
        PV_RunCvtXls = GetSettingBooleanParameter(GC_ENABLE_RUN_CVTXLS, False)
        UpdateActualProject
    End With
    
    ShowData "Form_Load"
    
    Set_MousePointer vbDefault
    
End Sub

Private Sub Form_LostFocus()

    GV_clsExportSpreadSheet.ClearLoaded

End Sub

Private Sub Form_Resize()

Dim lvGap           As Long

    With Me
        If .ScaleHeight > .txtSckData.Top Then
            .txtSckData.Height = .ScaleHeight - .txtSckData.Top - 120
        End If
        lvGap = 120
        If .ScaleWidth > 2 * PV_Width_Min Then
            .FrameResults.Width = .ScaleWidth - 2 * .FrameResults.Left
            .LstVwExport.Width = .FrameResults.Width - .LstVwExport.Left - lvGap
            .FrameActualProject.Width = .ScaleWidth - 2 * .FrameActualProject.Left
            .LstVwActualProject.Width = .FrameActualProject.Width - _
                                        2 * .LstVwActualProject.Left
        End If
        If .ScaleHeight > 2 * PV_Height_Min Then
            .FrameActualProject.Height = .ScaleHeight - .FrameActualProject.Top - lvGap
            .LstVwActualProject.Height = .FrameActualProject.Height - _
                                        .LstVwActualProject.Top - lvGap
        End If
        .txtSckData.Left = lvGap
        .txtSckData.Width = .pictureDeprecated.Width - 2 * lvGap
        .txtSckData.Top = lvGap
        .txtSckData.Height = .pictureDeprecated.Height - 2 * lvGap
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GV_clsExportSpreadSheet.ClearLoaded
    'GV_Mdi.Restore_Visible_MnuProject GenOutput
    GV_Mdi.ProjectMnuUpdate
    
    If PV_hFile Then
        Close #PV_hFile
        PV_hFile = 0
    End If
    
End Sub


Function NewFileName(lvStr As String, Index As Long)

Dim lvFormat            As String

    lvFormat = "00000" & Trim$(Str(Index))
    NewFileName = Left$(lvStr, Len(lvStr) - 4) & _
                    Right$(lvFormat, 5) & _
                    Right$(lvStr, 4)
    
End Function

Function SetArguments(lsFileName As String, SheetCount As Long) As String

Dim i           As Long

    For i = 0 To SheetCount - 1
        If i Then
            'SetArguments =SetSheetName(lsfilename)
        Else
        End If
    Next
    
End Function

Function GetMapFileName() As String

    GetMapFileName = m_Project.GetOutputPath & "\File.map"
    
End Function

Sub Create_Xls_File_Mark(lvName As String, lsArg As String)

Dim h               As Integer
Dim lFileName       As String
Dim lvFile          As String

   ''On Error GoTo Create_Xls_File_Mark_Error

    lFileName = Left$(lvName, Len(lvName) - 3) & "mrk"
    h = FreeFile
    Open lFileName For Output As h
        
    Print #h, lsArg
    
    Close #h

   'On Error GoTo 0
   Exit Sub

Create_Xls_File_Mark_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Create_Xls_File_Mark of Formulario frmExportSpreadSheet"
    Resume Next
    
End Sub

Function Verify_Exist_Map_File() As Boolean

Dim h               As Integer
Dim lFileName       As String

    On Error GoTo Verify_Exist_Map_File_Error

    Verify_Exist_Map_File = False
    lFileName = GetMapFileName

    If FileLen(lFileName) > 0 Then
        Verify_Exist_Map_File = True
    End If
    
    On Error GoTo 0
    Exit Function

Verify_Exist_Map_File_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Verify_Exist_Map_File of Formulario frmExportSpreadSheet"

End Function

Sub Create_Map_File()

Dim h               As Integer
Dim lFileName       As String
Dim i               As Long
Dim lCount          As Long
Dim lvFile          As String
Dim lsArguments     As String
Dim lvSheetCount    As Integer

   'On Error GoTo Create_Map_File_Error

    lFileName = GetMapFileName
    h = FreeFile
    
    Open lFileName For Output As h
        
    lCount = Pulse_GetSpreadFileCount
    
    Print #h, lCount
    For i = 0 To lCount - 1
        lvFile = NewFileName(PV_FileName, i)
        lvSheetCount = Pulse_GetSheetCount(i)
        lsArguments = Trim$(lvSheetCount) & "," & lvFile
        Print #h, lsArguments
    Next
    
    Close #h
    
    'Test_Proccess_Map_File lFileName
    
   'On Error GoTo 0
   Exit Sub

Create_Map_File_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Create_Map_File of Formulario frmExportSpreadSheet"
    Resume Next
    
End Sub

Private Sub TimerCreating_Timer()

Dim IndexFile       As Integer
Dim Pls             As Long
Dim lvNewName       As String
Dim lvDone          As Boolean
Dim lvXlsState      As Long
Dim lvXlsSaved      As Long
Dim lvSheetCount    As Long
Dim lvPulsesCount   As Long
Dim lvIndexSheet    As Long
Dim lvIndexPulse    As Long
Dim lvEnlased       As Date
Dim lsArguments     As String
Dim lsExe           As String
Dim lsOldState      As Integer

    'On Error GoTo timerCreating_Error
    
    lvEnlased = Now() - PV_Ini_Time
    Me.lblEnlasedTime.Caption = Format(lvEnlased, "hh:mm:ss")
    lsOldState = PV_State
    ShowData "TimerCreating_Timer" & ".  Old_State = " & PV_State
    Select Case PV_State
        Case Is = 0
            If PV_IndexSpread = 0 Then
                If PV_RunCvtXls = True Then
                    Create_Map_File
                'If GetSettingBooleanParameter(GC_ENABLE_RUN_CVTXLS, False) = False Then
                    'On Error GoTo Shell_Error
                    lsArguments = "1" & " " & GetMapFileName
                    lsExe = GV_XlsDll_Path & "\" & GV_Bin2Xls_App & " " & lsArguments
                    If GetSettingBooleanParameter(GC_ENABLE_HIDE_CVTXLS, False) = True Then
                        ShowData lsExe
                        Shell lsExe, vbHide
                    Else
                        ShowData lsExe
                        Shell lsExe
                    End If
                    'On Error GoTo timerCreating_Error
                End If
            End If
            lvNewName = NewFileName(PV_FileName, PV_IndexSpread)
            If GetSettingBooleanParameter(GC_ENABLE_XLS_OP, False) = False Then
                Pulse_Create_Xls_File lvNewName, PV_IndexSpread, PV_RunCvtXls
            Else
                Pulse_Create_Xls_File_Op lvNewName, PV_IndexSpread
            End If
            lvSheetCount = Pulse_GetSheetCount(PV_IndexSpread)
            With Me.ProgressBarAvance(1)
                .Value = 0
                .Max = lvSheetCount
            End With
            Me.ProgressBarAvance(2).Value = 0
            Me.SendXlsFileName lvNewName, PV_IndexSpread, lvSheetCount
            PV_State = 1
            ShowData "TimerCreating_Timer" & ".  New_State = " & PV_State
            'modLog.WriteLogFile "File: " & lvNewName
        Case Is = 1
            lvXlsState = Pulse_SpreadSheetDone(lvDone)
            lvXlsSaved = Pulse_SpreadSheet_Saved(lvDone)
            If lvXlsState And lvXlsSaved Then
                'modLog.WriteLogFile "File Saved: " & Me.lblEnlasedTime.Caption
                PV_State = 2
                ShowData "TimerCreating_Timer" & ".  New_State = " & PV_State
                lvNewName = NewFileName(PV_FileName, PV_IndexSpread)
                lvSheetCount = Pulse_GetSheetCount(PV_IndexSpread)
                lsArguments = Trim$(lvSheetCount) & " " & lvNewName
                'If PV_IndexSpread = 0 And _
                '    GetSettingBooleanParameter(GC_ENABLE_RUN_CVTXLS, False) = False Then
                '    'On Error GoTo Shell_Error
                '    lsArguments = "1" & " " & GetMapFileName
                '    lsExe = GV_XlsDll_Path & "\" & GV_Bin2Xls_App & " " & lsArguments
                '    If GetSettingBooleanParameter(GC_ENABLE_HIDE_CVTXLS, False) = True Then
                '        ShowData lsExe
                '        Shell lsExe, vbHide
                '    Else
                '        ShowData lsExe
                '        Shell lsExe
                '    End If
                '    'On Error GoTo timerCreating_Error
                '    'PV_State = 3
                'End If
                If PV_RunCvtXls = True Then
                    Create_Xls_File_Mark lvNewName, lsArguments
                    Me.SendXlsSaving PV_IndexSpread
                Else
                    Me.SendXlsFileReady PV_IndexSpread
                End If
            Else
                Pulse_SpreadSheetStatus lvIndexSheet, lvIndexPulse
                lvSheetCount = Pulse_GetSheetCount(PV_IndexSpread)
                If lvIndexSheet >= 0 Then
                    If Me.ProgressBarAvance(1).Value > lvIndexSheet Then
                        MsgBox "StatusBar greater than FileIndex.", vbCritical
                    End If
                    Me.ProgressBarAvance(1).Value = lvIndexSheet
                    ShowData "TimerCreating_Timer" & ".  Same_State = " & PV_State
                    If lvIndexSheet < lvSheetCount Then
                        lvPulsesCount = Pulse_GetSheetPulseCount(PV_IndexSpread, lvIndexSheet)
                        With Me.ProgressBarAvance(2)
                            .Max = lvPulsesCount
                            .Value = lvIndexPulse
                        End With
                        Me.SendXlsProgress PV_IndexSpread, lvIndexSheet, lvPulsesCount
                    Else
                        'modLog.WriteLogFile "Sheets Done: " & Me.lblEnlasedTime.Caption
                        With Me.ProgressBarAvance(2)
                            .Value = .Max
                        End With
                        lvPulsesCount = Pulse_GetSheetPulseCount(PV_IndexSpread, lvSheetCount - 1)
                        Me.SendXlsProgress PV_IndexSpread, lvSheetCount - 1, lvPulsesCount - 1
                        Me.SendXlsSaving PV_IndexSpread
                    End If
                Else
                    If PV_RunCvtXls = False And lvXlsState Then
                        Me.SendXlsSaving PV_IndexSpread
                    End If
                    'MsgBox "SpreadSheetStatus Error", vbCritical
                End If
            End If
        Case Is = 2
            PV_IndexSpread = PV_IndexSpread + 1
            Me.ProgressBarAvance(0).Value = PV_IndexSpread
            If PV_IndexSpread < Pulse_GetSpreadFileCount Then
                PV_State = 0
            ShowData "TimerCreating_Timer" & ".  New_State = " & PV_State
            Else
                If PV_RunCvtXls = True And Verify_Exist_Map_File = True Then
                    PV_IndexSpread = PV_IndexSpread - 1
                Else
                    Me.TimerCreating.Enabled = False
                    Me.cmdAccept.Enabled = True
                    Me.cmdCancel.Enabled = True
                ShowData "TimerCreating_Timer" & ".  Ending = " & PV_State
                    Me.EndingCreation
                    If GetSettingBooleanParameter(GC_ENABLE_XLS_OP, False) = True Then
                        Pulse_Finish_Xls
                    End If
                End If
            End If
        Case Is = 3
            If Find_App("cvt2xls.exe") = True Then
                PV_State = 4
            End If
        Case Is = 4
            If Find_App("cvt2xls.exe") = False Then
                PV_State = 5
            End If
        Case Is = 5
            Me.SendXlsFileReady PV_IndexSpread
            lvNewName = NewFileName(PV_FileName, PV_IndexSpread)
            lvSheetCount = Pulse_GetSheetCount(PV_IndexSpread)
            RemoveBinFiles lvSheetCount, lvNewName
            PV_State = 2
    End Select
    
    Exit Sub
    
Shell_Error:

    If GetSettingBooleanParameter(GC_ENABLE_HIDE_CVTXLS, False) = True Then
        MsgBox "Falla al Ejecutar:" & GV_XlsDll_Path & "\" & GV_Bin2Xls_App & " " & lsArguments, vbOKOnly
    Else
        MsgBox "Falla al Ejecutar:" & GV_XlsDll_Path & "\" & GV_Bin2Xls_App & " " & lsArguments, vbOKOnly
    End If
    Resume Next
    Exit Sub
    
timerCreating_Error:

    If PV_State = 4 And PV_State = lsOldState Then
        PV_State = 3
    Else
    MsgBox "Old State: " & lsOldState & "   New State: " & PV_State, vbOKOnly
    MsgBox Err.Description, vbOK
    End If
    Resume Next
    
End Sub

Sub RemoveBinFiles(lSheetCount As Long, lsXlsFile As String)

Dim i                   As Long
Dim lsBin               As String

    For i = 0 To lSheetCount - 1
        lsBin = GetBinFileName(lsXlsFile, i)
        'if getattr (lsxlsfile )=
        Kill lsBin
    Next
    
End Sub

Function GetBinFileName(lsXlsFile As String, Index As Long)

Dim lsNewName           As String
Dim lsBinName           As String
Dim midNewName          As String

    lsNewName = Left$(lsXlsFile, Len(lsXlsFile) - 4)
    midNewName = Mid$(lsNewName, Len(lsNewName) / 2)
    lsBinName = Right$("00" & Trim$(Index), 3) & ".bin"
    GetBinFileName = lsNewName & "_Sheet_" & lsBinName
    
End Function

Function GetArguments(lsFileName As String, lvSheetCount As Long)

Dim lsNewName           As String
Dim i                   As Long
Dim lsBinName           As String

    lsNewName = Left$(lsFileName, Len(lsFileName) - 4)
    For i = 1 To lvSheetCount
        lsBinName = Right$("00" & Trim$(i - 1), 3) & ".bin"
        If i = 1 Then
            GetArguments = lsNewName & lsBinName
        Else
            GetArguments = lsNewName & lsBinName
        End If
    Next
    
End Function

Sub TrVwExportItemSelected()

    'On Error GoTo Exit_Sub_TrVwExportItemSelected
    ShowData "TrVwExportItemSelected"
    If PV_TrVwIndex_Selected <> Me.TrVwExport.SelectedItem.Index Then
        PV_TrVwIndex_Selected = Me.TrVwExport.SelectedItem.Index
        modBgproject.ShowBgProject Me.LstVwExport, PV_TrVwIndex_Selected
    End If

Exit_Sub_TrVwExportItemSelected:

    'On Error GoTo 0
    
End Sub

Private Sub toolbarControles_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case Is = 1
            Me.cmdAccept.Value = True
            Me.LstVwActualProject.ListItems.Clear
            Button.Enabled = True
            m_Project.SaveProject
            'm_Project.CloseProject
            GV_Mdi.Close_Project False
            Button.Enabled = False
            'Me.Show
    End Select
    
End Sub

Private Sub TrVwExport_Click()

    Me.TrVwExportItemSelected
    
End Sub

Private Sub TrVwExport_KeyDown(KeyCode As Integer, Shift As Integer)

    Me.TrVwExportItemSelected
    
End Sub

Private Sub TrVwExport_KeyUp(KeyCode As Integer, Shift As Integer)

    Me.TrVwExportItemSelected
    
End Sub

Private Sub TrVwExport_NodeClick(ByVal Node As MSComctlLib.Node)


    Me.TrVwExportItemSelected
    
End Sub

