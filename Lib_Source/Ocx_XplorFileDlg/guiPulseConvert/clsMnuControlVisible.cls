VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMnuControlVisible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsMnuControlVisible
' Author    : Leo Herrera
' Date      : 14/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit


Private WithEvents mMnu            As Menu
Attribute mMnu.VB_VarHelpID = -1
Private mControl                   As Control

Public Sub SetControl(lvMnu As Menu, lvControl As Control)

    Set mMnu = lvMnu
    Set mControl = lvControl
    GetControlCheckedSetting mMnu
    
End Sub

Private Sub mMnu_Click()

Dim lvState         As Integer

    With mMnu
        If .Checked = True Then
            .Checked = False
        Else
            .Checked = True
        End If
        lvState = .Checked
        mControl.Visible = .Checked
        SaveSetting App.Title, _
                    GC_CONFIGURATION_SECTION, _
                    .Name & ".Checked", _
                    lvState
    End With
        
End Sub
