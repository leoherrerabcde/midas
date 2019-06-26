VERSION 5.00
Begin VB.Form frmStartUp 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   9615
   Begin VB.Timer tmrImageBtn 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pictureContainer 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   2
      Left            =   240
      ScaleHeight     =   2295
      ScaleWidth      =   6495
      TabIndex        =   2
      Top             =   5280
      Width           =   6495
      Begin VB.PictureBox PictureOpenPjtFromDisk 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   1335
         TabIndex        =   18
         Top             =   120
         Width           =   1335
         Begin VB.Image ImageBtn 
            Height          =   1095
            Index           =   2
            Left            =   120
            Picture         =   "frmStarUp.frx":0000
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "- Seleccionar Nueva Ubicación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   3615
      End
      Begin VB.Line line 
         Index           =   2
         X1              =   1560
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbl 
         Caption         =   "Proyectos Anteriores"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox pictureContainer 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   1
      Left            =   240
      ScaleHeight     =   2295
      ScaleWidth      =   6495
      TabIndex        =   1
      Top             =   2760
      Width           =   6495
      Begin VB.PictureBox PictureOpenPjt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   1335
         TabIndex        =   17
         Top             =   120
         Width           =   1335
         Begin VB.Image ImageBtn 
            Height          =   1095
            Index           =   1
            Left            =   120
            Picture         =   "frmStarUp.frx":41E8
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "ubicación en Disco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "- Abrir Proyecto de Conversión de Pulsos desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   4695
      End
      Begin VB.Line line 
         Index           =   1
         X1              =   1560
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbl 
         Caption         =   "Abrir Proyecto desde Disco"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox pictureContainer 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   0
      Left            =   240
      ScaleHeight     =   2295
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      Begin VB.PictureBox PictureNewPjt 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Index           =   0
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   1335
         TabIndex        =   16
         Top             =   240
         Width           =   1335
         Begin VB.Image ImageBtn 
            Height          =   1080
            Index           =   0
            Left            =   120
            Picture         =   "frmStarUp.frx":5C72
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1080
         End
      End
      Begin VB.PictureBox pictureNewProject 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1560
         ScaleHeight     =   375
         ScaleWidth      =   2535
         TabIndex        =   11
         Top             =   240
         Width           =   2535
         Begin VB.Label lbl 
            Caption         =   "Proyecto Nuevo"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.Label Label1 
         Caption         =   "- Configurar Salida"
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
         Index           =   2
         Left            =   1560
         TabIndex        =   5
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "- Seleccionar Ubicación Pulsos"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "- Seleccionar Ubicación Proyecto"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Line line 
         Index           =   0
         X1              =   1560
         X2              =   4560
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.PictureBox pictureShadowContainer 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   6960
      ScaleHeight     =   975
      ScaleWidth      =   1335
      TabIndex        =   15
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox pictureShadowContainer 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   6840
      ScaleHeight     =   975
      ScaleWidth      =   1335
      TabIndex        =   14
      Top             =   360
      Width           =   1335
   End
   Begin VB.PictureBox pictureShadowContainer 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   0
      Left            =   6720
      ScaleHeight     =   975
      ScaleWidth      =   1335
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmStartUp
' Author    : Leo Herrera
' Date      : 04/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private PC_Boton3D                          As clsBoton3D
Private mPictForm()         As clsDockFormToPict

Private Sub Close_Form()

    Unload Me
    
End Sub

Sub InitStartUpForm()

Dim i           As Integer

    Set PC_Boton3D = New clsBoton3D
    GV_Mdi.Set_Visible_Mnu_For_Open False
    With Me
        .WindowState = vbMaximized
        If m_Project.IsThereProjectList = False Then
            .pictureContainer(2).Enabled = False
        End If
        For i = .pictureContainer.LBound To .pictureContainer.UBound
            SetSameBackColor .pictureContainer(i), Me
        Next
'        ReDim mPictForm(.pictureContainer.UBound)
'        For i = .pictureContainer.LBound To .pictureContainer.UBound
'            Set mPictForm(i) = New clsDockFormToPict
'        Next
'        mPictForm(0).PickForm .pictureContainer(0), frmLinkNewProject

'        For i = .pictureContainer.LBound To .pictureContainer.UBound
'            m_MakeRound.MakeRoundControl .pictureContainer(i)
        'dockFormAndRound .pictureContainer(0), frmLinkNewProject
'        With .pictureContainer(0)
'            MakeRoundRect .hWnd, _
'                          .Width \ Screen.TwipsPerPixelX, _
'                        .Height \ Screen.TwipsPerPixelY, _
'                         20
'        End With
'        Next
    End With
    
End Sub

Private Sub Form_Load()

    With Me
        .InitStartUpForm
    End With
    Set_MousePointer vbDefault
    
End Sub

Private Sub Form_Resize()

Dim lv_Hori_Offset          As Long
Dim lv_Vert_Offset          As Long
Dim LV_Width                As Long
Dim lv_Height               As Long
Dim i                       As Integer

    With Me
        'Form_Move_Controls_To_Center Me, "picture"
'        Resize_Control .pictureLeftBar, Me, _
'                        refoOperationsConstants.refoVertical + refoOperationsConstants.refoToBorder, _
'                        .pictureLeftBar.Top
        Form_Resize_Controls Me, Me.pictureContainer(0).Name
        'mPictForm(0).Resize_Pict
        
'        For i = .pictureContainer.LBound To .pictureContainer.UBound
'            m_MakeRound.MakeRoundControl .pictureContainer(i)
'        Next

'        SetShadow .pictureLeftBar, .pictureShadowLeftBar, 20
'        m_MakeRound.MakeRoundControl .pictureLeftBar
'        m_MakeRound.MakeRoundControl .pictureShadowLeftBar
        For i = .pictureContainer.LBound To .pictureContainer.UBound
            SetShadow .pictureContainer(i), .pictureShadowContainer(i), 20
            m_MakeRound.MakeRoundControl .pictureShadowContainer(i)
            m_MakeRound.MakeRoundControl .pictureContainer(i)
        Next
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    GV_clsStartUp.ClearLoaded
    'mPictForm(0).Close_Form
    'GV_Mdi.Set_Visible_Mnu_For_Open True

End Sub

Private Sub ImageBtn_Click(Index As Integer)

    Select Case Index
        Case Is = 0
            GV_Mdi.NewProject
        Case Is = 1
            GV_Mdi.ProjectFromFile
        Case Is = 2
            GV_Mdi.ProjectFromListFile
    End Select
    
    Close_Form

End Sub

Private Sub ImageBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    PC_Boton3D.Iniciar Me.ImageBtn(Index), Me.ImageBtn(Index).Container, Me.tmrImageBtn

End Sub

