VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreprocessing 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAnalizingProccess 
      Interval        =   500
      Left            =   5040
      Top             =   4320
   End
   Begin MSComctlLib.ProgressBar ProgressBarAnalizingProccess 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lbl 
      Caption         =   "Archivos:"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblWrittenFiles 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
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
      Left            =   3720
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lbl 
      Caption         =   "/"
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
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label lblFilesCount 
      Caption         =   "100"
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
      Left            =   5040
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblDone 
      Alignment       =   2  'Center
      Caption         =   "Análisis Terminado"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblMisionName 
      Alignment       =   2  'Center
      Caption         =   "lblMisionName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Pre Procesando Pulsos Misión"
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6975
   End
End
Attribute VB_Name = "frmPreprocessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmPreprocessing
' Author    : Leo Herrera
' Date      : 16/07/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit




Private Sub tmrAnalizingProccess_Timer()

Dim IndexFilePwdLst, IndexFile, FileCount   As Long
Dim ProccessDone                            As Long

    With Me
        Pulse_Import_File_Status IndexFilePwdLst, _
                                    IndexFile, _
                                    FileCount, _
                                    ProccessDone
         
        WriteLogFile "IndexFilePwdLst = " & IndexFilePwdLst
        WriteLogFile "IndexFile = " & IndexFile
        WriteLogFile "FileCount = " & FileCount
        WriteLogFile "ProccessDone = " & ProccessDone
        
        .lblWrittenFiles.Caption = IndexFile
        .lblFilesCount.Caption = FileCount
        If FileCount Then
            With .ProgressBarAnalizingProccess
                .Max = FileCount
                .Value = IndexFile
            End With
        End If
        DoEvents
        Me.Refresh
        If ProccessDone Then
            WriteLogFile "Unload Me"
            Unload Me
            Exit Sub
        End If
    End With
    
End Sub
Private Sub Form_Load()

    Set_MousePointer vbHourglass
    With Me
        .Enabled = False
        .lblMisionName.Caption = m_Project.GetMissionName
        .Caption = m_Project.GetName
        WriteLogFile "Form_Load::m_Project.Run_Pulse_Analize_BG"
        m_Project.Run_Pulse_Analize_BG
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set_MousePointer vbDefault
    
End Sub
