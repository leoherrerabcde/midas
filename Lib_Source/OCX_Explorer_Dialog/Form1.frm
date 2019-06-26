VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10605
   FillColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   4680
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   10575
      TabIndex        =   1
      Top             =   720
      Width           =   10575
      Begin VB.Frame Frame1 
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   10335
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   3360
            Width           =   2415
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3015
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cambiar skin"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   3360
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.ImageList ImageListNormal 
      Left            =   7680
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":01C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":054F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imagelistup 
      Left            =   8280
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1020
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   1217
      ButtonWidth     =   1640
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageListNormal"
      HotImageList    =   "imagelistup"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "  Subir nivel"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Iconos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Lista"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Detalle"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Iconos pequeños"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   120
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   720
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Archivo"
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSubLV As cSubclassListView
Private WithEvents cFile As clsListFile
Attribute cFile.VB_VarHelpID = -1

Private Sub cFile_changePath(Ruta As String)
    Frame1.Caption = Ruta
End Sub

Private Sub Combo1_Click()
    Dim SpathSkin As String
    
    SpathSkin = App.Path & "\"
    
    Select Case Combo1.ListIndex
        Case 0: Call setColumnHeader(SpathSkin & Combo1.Text, vbBlack, vbBlack, True, False)
        Case 1: Call setColumnHeader(SpathSkin & Combo1.Text, vbBlack, vbBlack, True, False)
        Case 2: Call setColumnHeader(SpathSkin & Combo1.Text, vbBlack, vbBlack, True, False)
        Case 3: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, vbGreen, True, False)
        Case 4: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, vbBlack, True, False)
        Case 5: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, vbBlack, True, False)
        Case 6: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, vbBlack, True, False)
        Case 7: Call setColumnHeader(SpathSkin & Combo1.Text, &HC0FFC0, vbBlack, True, False)
        Case 8: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, vbBlack, True, False)
        Case 9: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, vbBlack, True, False)
        Case 10: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, vbBlack, True, False)
        Case 11: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, vbBlack, True, False)
        Case 12: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, vbYellow, True, False)
        Case 13: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, vbBlack, True, False)
        Case 14: Call setColumnHeader(SpathSkin & Combo1.Text, vbBlack, vbBlack, True, False)
        Case 15: Call setColumnHeader(SpathSkin & Combo1.Text, vbBlack, vbBlack, True, False)
        Case 16: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, vbBlack, True, False)
        Case 17: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, &H808080, True, False)
        Case 18: Call setColumnHeader(SpathSkin & Combo1.Text, vbYellow, vbWhite, True, False)
        Case 19: Call setColumnHeader(SpathSkin & Combo1.Text, vbBlack, vbBlack, True, False)
        Case 20: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, &H808080, True, False)
        Case 21: Call setColumnHeader(SpathSkin & Combo1.Text, vbWhite, &HC0FFFF, True, False)
        Case 22: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, &H808080, True, False)
        Case 23: Call setColumnHeader(SpathSkin & Combo1.Text, &HC0FFFF, &H808080, True, False)
        Case 24: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, vbWhite, True, False)
        Case 25: Call setColumnHeader(SpathSkin & Combo1.Text, &H808080, &H808080, True, False)
        
        
        
    End Select
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    Call InitCommonControls

    Call SetErrorMode(2)
End Sub

Private Sub Form_Load()

    Set cSubLV = New cSubclassListView
    Set cFile = New clsListFile
    
    With cFile
        .SetControls ListView1, ImageList1(0), ImageList1(1)
        .Listar "c:\"
        
    End With
    
    cSubLV.SubClassListView ListView1.hwnd
    
    Dim i As Integer
    
    For i = 1 To 26
        Combo1.AddItem "Skin" & i
    Next
    
    Combo1.ListIndex = 17
    


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cSubLV = Nothing
    Set cFile = Nothing
End Sub


Sub setColumnHeader( _
    SpathSkin As String, _
    lColorNormal As Long, _
    lColorUp As Long, _
    Optional bIconAlingmentRight As Boolean = False, _
    Optional bTextBold As Boolean = False)
    
    With cSubLV
        .SkinPicture = LoadPicture(SpathSkin & ".bmp")
        .TextNormalColor = lColorNormal
        .TextResalteColor = lColorUp
        .IconAlingmentRight = bIconAlingmentRight
        .HedersFontBlod = bTextBold
    End With
End Sub


Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: cFile.subirNivel
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    Select Case ButtonMenu.Index
        Case 1: ListView1.View = lvwIcon
        Case 2: ListView1.View = lvwList
        Case 3: ListView1.View = lvwReport
        Case 4: ListView1.View = lvwSmallIcon: ListView1.Arrange = lvwAutoLeft
    End Select
End Sub
