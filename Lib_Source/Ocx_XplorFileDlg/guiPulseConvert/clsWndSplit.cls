VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWndSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents imgSplitter          As Image
Attribute imgSplitter.VB_VarHelpID = -1
Private WithEvents picSplitter          As PictureBox
Attribute picSplitter.VB_VarHelpID = -1
Private WithEvents LeftList             As ListView
Attribute LeftList.VB_VarHelpID = -1
Private WithEvents RightList            As ListView
Attribute RightList.VB_VarHelpID = -1
Private WithEvents FrameContainer       As Frame
Attribute FrameContainer.VB_VarHelpID = -1
Private objIzq                          As ListView
Private objDer                          As ListView
Private FormName                        As String
'Private Left_Pict                       As PictureBox
Private Right_Pict                      As PictureBox
Private PV_Min_Width                    As Long
Private FormContainer                   As Form

Private moviendo As Boolean

Sub Constructor(LV_Image As Image, LV_picture As PictureBox, LV_Frame As Frame, _
                LV_L As ListView, LV_R As ListView, LV_FrmContainer As Form, _
                LV_R_Pict As PictureBox, Optional LV_Min_Width As Long)

    Set imgSplitter = LV_Image
    Set picSplitter = LV_picture
    Set FrameContainer = LV_Frame
    Set objIzq = LV_L
    Set objDer = LV_R
    Set FormContainer = LV_FrmContainer
    Set Right_Pict = LV_R_Pict
    If LV_Min_Width Then
        PV_Min_Width = LV_Min_Width
    Else
        PV_Min_Width = 90
    End If
    
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 3, .Height - 20
    End With
    picSplitter.Visible = True
    moviendo = True

End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim sglPos As Single
    '
    
    If moviendo Then
        sglPos = X + imgSplitter.Left
        If sglPos < splitLimit Then
            picSplitter.Left = splitLimit
        ElseIf sglPos > FormContainer.Width - splitLimit Then
            picSplitter.Left = FormContainer.Width - splitLimit
        Else
            picSplitter.Left = sglPos
        End If
    Else
    End If

End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With Me
        sizeControls picSplitter.Left
        picSplitter.Visible = False
    End With
    moviendo = False

End Sub

Sub sizeControls(ByVal X As Long)
    
Dim tMinWidth As Long
    '
    On Error Resume Next
    '
    ' el ancho mínimo que tendrá cada panel
    tMinWidth = Screen.TwipsPerPixelY * PV_Min_Width
    '
    ' asignar el ancho
    If X < tMinWidth Then X = tMinWidth
    If X > (FrameContainer.Width - tMinWidth) Then X = FrameContainer.Width - tMinWidth
    objIzq.Width = X - objIzq.Left - imgSplitter.Width
    imgSplitter.Left = X
    SaveSetting App.Title, _
                LV_Name, _
                imgSplitter.Name & ".Left", _
                imgSplitter.Left
    objDer.Left = X + imgSplitter.Width
    objDer.Width = FrameContainer.Width - objDer.Left - objIzq.Left '140)
    
    Right_Pict.Left = X
    'Me.toolbarControles(1).Left = Me.lblPreView.Left + Me.lblPreView.Width
    
    imgSplitter.Top = objIzq.Top
    imgSplitter.Height = objIzq.Height

End Sub



