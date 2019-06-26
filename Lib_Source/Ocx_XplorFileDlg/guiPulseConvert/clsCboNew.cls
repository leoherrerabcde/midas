VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCboNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents m_CboBox         As ComboBox
Attribute m_CboBox.VB_VarHelpID = -1
Private WithEvents m_TxtBox         As TextBox
Attribute m_TxtBox.VB_VarHelpID = -1
Private WithEvents m_cmdAccept      As CommandButton
Attribute m_cmdAccept.VB_VarHelpID = -1
Private WithEvents m_cmdCancel      As CommandButton
Attribute m_cmdCancel.VB_VarHelpID = -1
Private WithEvents m_Form           As Form
Attribute m_Form.VB_VarHelpID = -1

Private m_clsSelText                As clsSelText
Private m_LblWarning                As Label

Private m_StrNew                    As String
Private m_IndexNew                  As Integer

Public Event AddNewItem()
Public Event NewItemAdded()
Public Event OnClickNewItem(Index As Integer)

Sub AddItem_New(lsText As String)

    m_IndexNew = m_CboBox.ListCount
    m_StrNew = lsText
    m_TxtBox.Text = m_StrNew
    m_CboBox.AddItem lsText
    SetVisibleCbo True
    SetVisibleWarn False
    
End Sub

Private Sub CloseNewItem()

    SetVisibleCbo True
    SetVisibleWarn False
    m_TxtBox.Text = m_StrNew
    
End Sub

Private Sub FlagNotUniqueNameWarning(lvVisible As Boolean)

    m_LblWarning.Visible = lvVisible
    
End Sub

Private Sub AddItem_TxtBox()

    m_CboBox.RemoveItem m_IndexNew
    m_CboBox.AddItem m_TxtBox.Text
    CloseNewItem
    AddItem_New m_StrNew

End Sub

Private Sub ChangeItemText()

    m_CboBox.List(m_CboBox.ListIndex) = m_TxtBox.Text
    
End Sub
Private Sub NewItem()

Dim Index           As Integer

    If VerifyUniqueName = True Then
        SetVisibleWarn False
        'AddItem_TxtBox
        ChangeItemText
        CloseNewItem
        RaiseEvent NewItemAdded
        AddItem_New m_StrNew
        'm_CboBox.ListIndex = m_CboBox.ListCount - 2
        Index = m_CboBox.ListCount - 2
        If Index < 0 Then
            Index = 0
        End If
        RaiseEvent OnClickNewItem(Index)
    End If
    
End Sub

Sub SetControls(Cbo As ComboBox, LV_Form As Form, Txt As TextBox, cmdAccept As CommandButton, _
                cmdCancel As CommandButton, LblWarn As Label, _
                Optional AddItemNew As Boolean)

    Set m_CboBox = Cbo
    Set m_TxtBox = Txt
    Set m_cmdAccept = cmdAccept
    Set m_cmdCancel = cmdCancel
    Set m_LblWarning = LblWarn
    Set m_clsSelText = New clsSelText
    Set m_Form = LV_Form
    
    m_clsSelText.SetControl Txt
    m_TxtBox.Top = m_CboBox.Top
    m_TxtBox.Left = m_CboBox.Left
    SetVisibleWarn False
    SetVisibleCbo True
    If AddItemNew = True Then
        Me.AddItem_New "Nuevo..."
    End If
    
End Sub

Sub SetVisibleCbo(lsVal As Boolean)

    If lsVal = True Then
        m_CboBox.Visible = True
        m_TxtBox.Visible = False
        m_cmdAccept.Visible = False
        m_cmdCancel.Visible = False
    Else
        m_CboBox.Visible = False
        m_TxtBox.Visible = True
        m_cmdAccept.Visible = True
        m_cmdCancel.Visible = True
    End If
    
End Sub

Private Sub SetVisibleWarn(lvVisible As Boolean)
    
        m_LblWarning.Visible = lvVisible

End Sub

Private Sub VerifyNewItem()

    If m_CboBox.ListIndex = m_IndexNew Then
        SetVisibleCbo False
'        Load m_Form
'        m_Form.Show vbModal
        RaiseEvent AddNewItem
    End If

End Sub

Private Function VerifyUniqueName() As Boolean

Dim i           As Integer

    VerifyUniqueName = False
    With m_CboBox
        For i = 0 To .ListCount - 2
            If .List(i) = m_TxtBox Then
                FlagNotUniqueNameWarning True
                Exit Function
            End If
        Next
    End With
    FlagNotUniqueNameWarning False
    VerifyUniqueName = True
    
End Function

Private Sub m_CboBox_Change()

    VerifyNewItem
    
End Sub

Private Sub m_CboBox_Click()

    VerifyNewItem

End Sub

Private Sub m_CboBox_KeyDown(KeyCode As Integer, Shift As Integer)

    VerifyNewItem

End Sub

Private Sub m_CboBox_KeyUp(KeyCode As Integer, Shift As Integer)

    VerifyNewItem

End Sub

Private Sub m_cmdAccept_Click()

    NewItem
    
End Sub

Private Sub m_cmdCancel_Click()

    CloseNewItem
    
End Sub

Private Sub m_Form_Unload(Cancel As Integer)

    m_Text = m_Form.m_Text
    
End Sub

'Private Sub m_CboBox_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'        NewItem
'    End If
'
'End Sub
Private Sub m_TxtBox_Change()

End Sub

Private Sub m_TxtBox_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        NewItem
    End If

End Sub
