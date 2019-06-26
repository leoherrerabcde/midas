VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTxtCboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mTextBox As TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mComboBox As ComboBox
Attribute mComboBox.VB_VarHelpID = -1

Public Event OnLostFocus()


Private PV_Init_Cbo_Txt         As Boolean

'Private WithEvents mDtPicker As DTPicker

Function Cbo_Box_Visible() As Boolean

    Cbo_Box_Visible = mComboBox.Visible
    
End Function

Sub Traspasar_Data_To_Text()
    
    mTextBox.Text = mComboBox.Text
    PV_Init_Cbo_Txt = True
    mComboBox.Visible = False
    
End Sub

Sub Iniciar(LV_TextBox As TextBox, LV_Form As Form, Optional LV_Container As Control)

    If mTextBox Is Nothing Then
        Set mTextBox = LV_TextBox
    End If
    
    Set mComboBox = LV_Form.Controls.Add("vb.combobox", "cbo_" & mTextBox.Name)
    
    If LV_Container Is Nothing Then
    Else
        Set mComboBox.Container = LV_Container
    End If
    
    mComboBox.Left = mTextBox.Left
    'mComboBox.Height = mTextBox.Height
    mComboBox.Width = mTextBox.Width
    mComboBox.Top = mTextBox.Top
    
    mComboBox.Visible = False
    PV_Init_Cbo_Txt = True
    
End Sub

Sub AddItem(Item As String)

    If mComboBox Is Nothing Then
        Exit Sub
    End If
    
    mComboBox.AddItem Item
    
End Sub

Sub Clear()

    If mComboBox Is Nothing Then
        Exit Sub
    End If
    
    mComboBox.Clear
    
End Sub

Sub Set_Visible_Cbo()

Dim i           As Integer

    If PV_Init_Cbo_Txt = True Then
        PV_Init_Cbo_Txt = False
        mComboBox.ListIndex = 0
        For i = 0 To mComboBox.ListCount - 1
            If mTextBox.Text = mComboBox.List(i) Then
                mComboBox.ListIndex = i
                Exit For
            End If
        Next
    End If
    mComboBox.Visible = True
    
End Sub

Private Sub mComboBox_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        
    End If
    
End Sub

Private Sub mComboBox_LostFocus()

    Traspasar_Data_To_Text
    
End Sub

Private Sub mTextBox_Click()

    Set_Visible_Cbo
    
End Sub

Private Sub mTextBox_GotFocus()

    Set_Visible_Cbo
    
End Sub

Private Sub mTextBox_KeyPress(KeyAscii As Integer)

    Set_Visible_Cbo
    
End Sub
