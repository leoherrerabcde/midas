VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSelText
' Author    : Leo Herrera
' Date      : 12/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private WithEvents mTextBox        As TextBox
Attribute mTextBox.VB_VarHelpID = -1

Sub SetControl(lvTxtBx As TextBox)

    Set mTextBox = lvTxtBx
    
End Sub

Private Sub mTextBox_GotFocus()

    With mTextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub
