VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDialogAnalizingData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2460
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnalizingProccess 
      Interval        =   200
      Left            =   2880
      Top             =   2040
   End
   Begin MSComctlLib.ProgressBar ProgressBarAnalizingProccess 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   495
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
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   135
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
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   375
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
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmDialogAnalizingData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub tmrAnalizingProccess_Timer()

Dim IndexFilePwdLst, IndexFile, FileCount   As Long
Dim ProccessDone                            As Long

    With Me
        Pulse_Import_File_Status IndexFilePwdLst, _
                                    IndexFile, _
                                    FileCount, _
                                    ProccessDone
        
        If ProccessDone Then
            Unload Me
        End If
        .lblWrittenFiles.Caption = IndexFile
        .lblFilesCount.Caption = FileCount
        If FileCount Then
            With .ProgressBarAnalizingProccess
                .Max = FileCount
                .Value = IndexFile
            End With
        End If
    End With
    
End Sub
