VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl XplorerDialog 
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ScaleHeight     =   5640
   ScaleWidth      =   5895
   Begin MSComctlLib.TreeView TrVwHightLight 
      Height          =   3615
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6376
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin VB.ComboBox cboSkin 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   "cboSkin"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView LstVwFile 
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7858
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageListNormal 
      Left            =   3720
      Top             =   4920
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
            Picture         =   "XplorerDialog.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XplorerDialog.ctx":01C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XplorerDialog.ctx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XplorerDialog.ctx":054F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imagelistup 
      Left            =   4320
      Top             =   4920
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
            Picture         =   "XplorerDialog.ctx":0AE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XplorerDialog.ctx":0D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XplorerDialog.ctx":1020
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "XplorerDialog.ctx":12BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   120
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   720
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label lbl 
      Caption         =   "888888"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "XplorerDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim cSubLV As cSubclassListView

Private WithEvents cFile As clsListFile
Attribute cFile.VB_VarHelpID = -1
Public Event OnCurrentFolderChanged(ByVal NewFolder As String)
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event OnItemClick(ByVal Item As MSComctlLib.ListItem)

Private PV_CurrentFolder
Private PV_ListHightLight() As String

Property Let HighLightPathForeColor(Color As Long)

    cFile.mHighLigthPathForeColor = Color
    
End Property

Property Let HighLightPathFullForeColor(Color As Long)

    cFile.mHighLigthPathFullBackColor = Color
    
End Property

Property Let HighLightPathBackColor(Color As Long)

    cFile.mHighLigthPathBackColor = Color
    cFile.mHighLigthPathBackColorEnable = True
    
End Property

Property Let HighLightPathFullBackColor(Color As Long)

    cFile.mHighLigthPathFullBackColor = Color
    cFile.mHighLigthPathBackColorEnable = True
    
End Property

Property Let HighLightMissionForeColor(Color As Long)

    cFile.mHighLigthMissionForeColor = Color
    
End Property

Property Let HighLightMissionBackColor(Color As Long)

    cFile.mHighLigthMissionBackColor = Color
    cFile.mHighLigthMissionBackColorEnable = True
    
End Property

'-------------------------------------------------------
Property Get HighLightPathForeColor() As Long

    HighLightPathForeColor = cFile.mHighLigthPathBackColor
    
End Property

Property Get HighLightPathFullForeColor() As Long

    HighLightPathFullForeColor = cFile.mHighLigthPathFullForeColor
    
End Property

Property Get HighLightPathBackColor() As Long

    HighLightPathBackColor = cFile.mHighLigthPathBackColor
    
End Property

Property Get HighLightPathFullBackColor() As Long

    HighLightPathBackColor = cFile.mHighLigthPathFullBackColor
    
End Property

Property Get HighLightMissionForeColor() As Long

    HighLightMissionForeColor = cFile.mHighLigthMissionForeColor
    
End Property

Property Get HighLightMissionBackColor() As Long

    HighLightMissionBackColor = cFile.mHighLigthMissionBackColor
    
End Property
'
'--------------------------------------------------------
Property Let AppTitle(lsTitle As String)

    GV_App_Title = lsTitle
    
End Property

Property Get Version() As String

    Version = App.Title & ":" & App.Major & "," & App.Minor & "," & App.Revision
    
End Property

Property Get SkinName() As String

    With UserControl.cboSkin
        If .ListIndex >= 0 And .ListIndex < .ListCount Then
            SkinName = .List(.ListIndex)
        Else
            SkinName = ""
        End If
    End With
    
End Property

Property Get SkinIndex() As Integer

    SkinIndex = UserControl.cboSkin.ListIndex
    
End Property

Property Get DebugTreeView() As Boolean

    With UserControl
        DebugTreeView = .TrVwHightLight.Visible
    End With

End Property

Property Let DebugTreeView(lvVisible As Boolean)

    With UserControl
        .TrVwHightLight.Visible = lvVisible
        If lvVisible = True Then
            .LstVwFile.Visible = False
        Else
            .LstVwFile.Visible = True
        End If
    End With
    
End Property

Property Get HightLight() As Boolean

    HightLight = cFile.HightLight
    
End Property

Property Let HightLight(lbValue As Boolean)

    cFile.HightLight = lbValue
    
End Property

Property Let SkinIndex(Index As Integer)

    With UserControl.cboSkin
        If Index >= -1 And Index < .ListCount Then
            .ListIndex = Index
        End If
    End With
    
End Property

Property Get CurrentFolder() As String

    CurrentFolder = PV_CurrentFolder
    
End Property

Property Let CurrentFolder(lvCurrentFolder As String)

    cFile.Listar lvCurrentFolder
    If PV_CurrentFolder <> lvCurrentFolder Then
        PV_CurrentFolder = lvCurrentFolder
        RaiseEvent OnCurrentFolderChanged(PV_CurrentFolder)
    End If
    
End Property

Property Get SelectedCount() As Long

Dim i       As Long
    
    SelectedCount = 0
    With UserControl.LstVwFile
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected = True Then
                SelectedCount = SelectedCount + 1
            End If
        Next
    End With
    
End Property

Property Get FirstSelectedItem() As ListItem

    Set FirstSelectedItem = UserControl.LstVwFile.SelectedItem
    
End Property

Property Get ViewStyle() As ListViewConstants

    ViewStyle = UserControl.LstVwFile.View
    
End Property

Property Let ViewStyle(lvView As ListViewConstants)

    UserControl.LstVwFile.View = lvView
    
End Property

Public Sub RefreshView()

    UserControl.Refresh
    cFile.Listar PV_CurrentFolder
    
End Sub

Private Sub cFile_changePath(Ruta As String)
    'Frame1.Caption = Ruta
    If PV_CurrentFolder <> Ruta Then
        PV_CurrentFolder = Ruta
        RaiseEvent OnCurrentFolderChanged(Ruta)
    End If
    
End Sub

Private Sub cboSkin_Click()

Dim SpathSkin       As String
Dim lsBmpPath       As String
    
    SpathSkin = App.Path
    SpathSkin = Obtener_Configuracion("Configuration", "SpathSkin", SpathSkin)
    lsBmpPath = Obtener_Configuracion("Configuration", "BmpPath", "Iconos")
    SpathSkin = FindPath(SpathSkin, lsBmpPath) & "\"
    
    
    Select Case cboSkin.ListIndex
        Case 0: Call setColumnHeader(SpathSkin & cboSkin.Text, vbBlack, vbBlack, True, False)
        Case 1: Call setColumnHeader(SpathSkin & cboSkin.Text, vbBlack, vbBlack, True, False)
        Case 2: Call setColumnHeader(SpathSkin & cboSkin.Text, vbBlack, vbBlack, True, False)
        Case 3: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, vbGreen, True, False)
        Case 4: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, vbBlack, True, False)
        Case 5: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, vbBlack, True, False)
        Case 6: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, vbBlack, True, False)
        Case 7: Call setColumnHeader(SpathSkin & cboSkin.Text, &HC0FFC0, vbBlack, True, False)
        Case 8: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, vbBlack, True, False)
        Case 9: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, vbBlack, True, False)
        Case 10: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, vbBlack, True, False)
        Case 11: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, vbBlack, True, False)
        Case 12: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, vbYellow, True, False)
        Case 13: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, vbBlack, True, False)
        Case 14: Call setColumnHeader(SpathSkin & cboSkin.Text, vbBlack, vbBlack, True, False)
        Case 15: Call setColumnHeader(SpathSkin & cboSkin.Text, vbBlack, vbBlack, True, False)
        Case 16: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, vbBlack, True, False)
        Case 17: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, &H808080, True, False)
        Case 18: Call setColumnHeader(SpathSkin & cboSkin.Text, vbYellow, vbWhite, True, False)
        Case 19: Call setColumnHeader(SpathSkin & cboSkin.Text, vbBlack, vbBlack, True, False)
        Case 20: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, &H808080, True, False)
        Case 21: Call setColumnHeader(SpathSkin & cboSkin.Text, vbWhite, &HC0FFFF, True, False)
        Case 22: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, &H808080, True, False)
        Case 23: Call setColumnHeader(SpathSkin & cboSkin.Text, &HC0FFFF, &H808080, True, False)
        Case 24: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, vbWhite, True, False)
        Case 25: Call setColumnHeader(SpathSkin & cboSkin.Text, &H808080, &H808080, True, False)
        
    End Select
End Sub

'Private Sub Command1_Click()
'    Unload Me
'End Sub
'


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
    
   On Error GoTo setColumnHeader_Error

    With cSubLV
        .SkinPicture = LoadPicture(SpathSkin & ".bmp")
        .TextNormalColor = lColorNormal
        .TextResalteColor = lColorUp
        .IconAlingmentRight = bIconAlingmentRight
        .HedersFontBlod = bTextBold
    End With

   On Error GoTo 0
   Exit Sub

setColumnHeader_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setColumnHeader of Control de usuario XplorerDialog"
End Sub

Public Sub SetHighLightList(lsListHighLight() As String)

'Dim iSize               As Integer
'Dim i                   As Integer
'
'    iSize = UBound(lsListHighLight)
'    ReDim PV_ListHightLight(iSize)
'    For i = 0 To iSize
'        PV_ListHightLight(i) = lsListHighLight(i)
'    Next
    
    cFile.SetListFileHightLight lsListHighLight
    
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
        Case 1: LstVwFile.View = lvwIcon
        Case 2: LstVwFile.View = lvwList
        Case 3: LstVwFile.View = lvwReport
        Case 4: LstVwFile.View = lvwSmallIcon: LstVwFile.Arrange = lvwAutoLeft
    End Select
End Sub


Private Sub LstVwFile_Click()

    RaiseEvent Click
    
End Sub

Private Sub LstVwFile_DblClick()

    RaiseEvent DblClick
    
End Sub

Private Sub LstVwFile_ItemClick(ByVal Item As MSComctlLib.ListItem)

    RaiseEvent OnItemClick(Item)
    
End Sub

Private Sub LstVwFile_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)
    
End Sub

Private Sub LstVwFile_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 13 Then
        With UserControl
            'If .LstVwFile.SelectedItem <> Nothing Then
            'End If
        End With
    End If
    
End Sub

Private Sub LstVwFile_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
    
End Sub

Private Sub UserControl_Initialize()

    GV_App_Title = App.Title
    
    Call InitCommonControls

    Call SetErrorMode(2)

    Set cSubLV = New cSubclassListView
    Set cFile = New clsListFile
    
    With cFile
        .SetControls LstVwFile, ImageList1(0), ImageList1(1), TrVwHightLight
        .Listar "c:\"
        
    End With
    
    cSubLV.SubClassListView LstVwFile.hwnd
    
    Dim i As Integer
    
    cboSkin.Clear
    
    For i = 1 To 26
        cboSkin.AddItem "Skin" & i
    Next
    
    cboSkin.ListIndex = 17

End Sub

Private Sub UserControl_Resize()

    With UserControl
        .LstVwFile.left = .ScaleLeft
        .LstVwFile.Top = .ScaleTop
        .LstVwFile.Width = .ScaleWidth
        .LstVwFile.Height = .ScaleHeight
        .TrVwHightLight.left = .ScaleLeft
        .TrVwHightLight.Top = .ScaleTop
        .TrVwHightLight.Width = .ScaleWidth
        .TrVwHightLight.Height = .ScaleHeight
    End With
    
End Sub

Private Sub UserControl_Terminate()

    Set cSubLV = Nothing
    Set cFile = Nothing

End Sub
