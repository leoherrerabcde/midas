VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "frmStart"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrImageBtn 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1800
      Top             =   3960
   End
   Begin VB.PictureBox PictureImgBtn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.PictureBox PictureLblPjtFromDisk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   1560
         ScaleHeight     =   1095
         ScaleWidth      =   3015
         TabIndex        =   6
         Top             =   2880
         Width           =   3015
         Begin VB.Label lblBoton 
            Caption         =   "Abrir un Proyecto desde Disco"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.PictureBox PictureLblOpenPjt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   1560
         ScaleHeight     =   855
         ScaleWidth      =   2775
         TabIndex        =   5
         Top             =   1560
         Width           =   2775
         Begin VB.Label lblBoton 
            Caption         =   "Abrir un Proyecto Anterior"
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
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.PictureBox PictureLblNewPjt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1560
         ScaleHeight     =   735
         ScaleWidth      =   2775
         TabIndex        =   4
         Top             =   360
         Width           =   2775
         Begin VB.Label lblBoton 
            Caption         =   "Crear Nuevo Proyecto"
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
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.PictureBox PictureOpenPjtFromDisk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   1335
         TabIndex        =   3
         Top             =   2880
         Width           =   1335
         Begin VB.Image ImageBtn 
            Height          =   1095
            Index           =   2
            Left            =   120
            Picture         =   "frmStart.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.PictureBox PictureOpenPjt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   1335
         TabIndex        =   2
         Top             =   1440
         Width           =   1335
         Begin VB.Image ImageBtn 
            Height          =   1095
            Index           =   1
            Left            =   120
            Picture         =   "frmStart.frx":41E8
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.PictureBox PictureNewPjt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1335
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         Begin VB.Image ImageBtn 
            Height          =   1080
            Index           =   0
            Left            =   120
            Picture         =   "frmStart.frx":5C72
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmStart
' Author    : Leo Herrera
' Date      : 16/07/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


'---------------------------------------------------------------------------------------
' Module    : frmStart
' Author    : Leo Herrera
' Date      : 25/03/2011
' Purpose   :
'---------------------------------------------------------------------------------------


Private PC_Boton3D                          As clsBoton3D

Private Sub Init_Form()

'Dim i           As Integer
'
'    ReDim PC_Boton3D(Me.ImageBtn.UBound)
'
'    For i = 0 To Me.ImageBtn.UBound
'        Set PC_Boton3D(i) = New clsBoton3D
'    Next
    Set PC_Boton3D = New clsBoton3D
    
    'GC_Log.Enumerate_Controls Me
    Me.Caption = "Start"
    
    Me.PictureImgBtn.BorderStyle = 0
    
End Sub

Private Sub Close_Form()

    Unload Me
    
End Sub

Private Sub Form_Load()

    Init_Form
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'GC_Log.Write_Frm_Event_NoArg Me, "-1", "UNLOAD"
    GV_clsStart.ClearLoaded
    
End Sub

Private Sub ImageBtn_Click(Index As Integer)


    'GC_Log.Write_Frm_Event_NoArg Me, Me.ImageBtn(Index).Tag, "CLICK"
    
    'GV_Mdi.PU_Mnu_Selected = Index
    Close_Form
'    Select Case Index
'        Case 0
'            GV_MDI.Shorcut_NewProject
'            Close_Form
'        Case 1
'            GV_MDI.Shorcut_Recient_Prj
'            Close_Form
'        Case 2
'            GV_MDI.Shorcut_OpenPrj_From_File
'            Close_Form
'    End Select
    
End Sub

Sub LoadNewProject()

    GV_clsProjectSelLocation.SetLoad
    
End Sub

Sub LoadProjectFromFile()
    GV_clsPjtFromFile.SetLoad
End Sub

Sub LoadProjectFormList()
    GV_clsPjtFromList.SetLoad
End Sub

Private Sub ImageBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'    PC_Boton3D(Index).Iniciar Me.ImageBtn(Index), Me.PictureImgBtn
    PC_Boton3D.Iniciar Me.ImageBtn(Index), Me.ImageBtn(Index).Container, Me.tmrImageBtn
    
End Sub

Private Sub lblBoton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    PC_Boton3D.Iniciar Me.lblBoton(Index), Me.lblBoton(Index).Container, Me.tmrImageBtn

End Sub

