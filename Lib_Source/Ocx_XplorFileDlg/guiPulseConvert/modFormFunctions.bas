Attribute VB_Name = "modFormFunctions"
'---------------------------------------------------------------------------------------
' Module    : modFormFunctions
' Author    : Leo Herrera
' Date      : 29/11/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Enum FormMoveControlAlignmentConstants
    ToTheRight = 1
    ToTheLeft = 2
    ToTheCenter = 2
    OffSet = 8
End Enum

Sub GetControlCheckedSetting(LV_Control As Control)

    With LV_Control
        .Checked = GetBooleanSetting(GC_CONFIGURATION_SECTION, .Name & ".Checked", .Checked)
    End With
    
End Sub

Sub SaveControlCheckedSetting(LV_Control As Control)

Dim lv_State            As Integer

    With LV_Control
        lv_State = 0
        If .Checked = True Then
            lv_State = 1
        End If
        SaveSetting App.Title, GC_CONFIGURATION_SECTION, .Name & ".Checked", lv_State
    End With

End Sub

Sub SetSameBackColor(LV_Control As Control, LV_Form As Form)

Dim LV_Item             As Control

    On Error Resume Next
    For Each LV_Item In LV_Form.Controls
        If LV_Item.Container Is LV_Control Then
            LV_Item.BackColor = LV_Control.BackColor
        End If
    Next
    'On Error GoTo 0
    
End Sub

Sub CopyBackGroundColor(LV_Form As Form)

Dim LV_Control      As Control
Dim LV_Lbl          As Label

    For Each LV_Control In LV_Form.Controls
        If Left$(LV_Control.Name, 3) = "lbl" Then
            Set LV_Lbl = LV_Control
            LV_Lbl.BackColor = LV_Form.BackColor
        End If
    Next
    
End Sub


Sub dockFormAndRound(LV_Pict As PictureBox, LV_Form As Form)

    dockFormPict LV_Form.hwnd, LV_Pict, True
'    m_MakeRound.MakeRoundForm LV_Form

    With LV_Form
        MakeRoundRect .hwnd, _
                      .Width \ Screen.TwipsPerPixelX, _
                    .Height \ Screen.TwipsPerPixelY, _
                     20
    End With

End Sub

Public Function ControlTypeFit(LV_Control As Control, lvCriteria As String) As Boolean

    ControlTypeFit = False
    If lvCriteria <> "" Then
        If LCase(Left$(LV_Control.Name, Len(lvCriteria))) = LCase(lvCriteria) Then
            ControlTypeFit = True
        End If
    Else
        ControlTypeFit = True
    End If
    
End Function

Public Function ControlNameFit(LV_Control As Control, lvCriteria As String) As Boolean

    ControlNameFit = False
    If lvCriteria <> "" Then
        If LV_Control.Name = lvCriteria Then
            ControlNameFit = True
        End If
    Else
        ControlNameFit = True
    End If
    
End Function

Public Function ControlHasPositionProperty(LV_Control As Control) As Boolean

    ControlHasPositionProperty = False
    'On Error GoTo ControlHasPositionProperty_Error

    With LV_Control
        If .Left >= 0 And _
            .Width >= 0 And _
             .Top >= 0 And _
             .Height >= 0 Then
            ControlHasPositionProperty = True
        End If
    End With

    'On Error GoTo 0
    Exit Function

ControlHasPositionProperty_Error:

End Function

Public Sub Form_Resize_Controls(LV_Form As Form, Optional ControlName As String)

Dim LV_Control              As Control

    For Each LV_Control In LV_Form.Controls
        If ControlNameFit(LV_Control, ControlName) = True Then
            Form_Resize_Control LV_Form, LV_Control
        End If
    Next
    
End Sub

Public Sub SetShadow(LV_Control As Control, LV_Shadow As Control, Optional LV_Width As Long)

    If LV_Width = 0 Then
        LV_Width = 1
    End If
    
    LV_Shadow.Left = LV_Control.Left - LV_Width
    LV_Shadow.Width = LV_Control.Width + 2 * LV_Width
    LV_Shadow.Top = LV_Control.Top - LV_Width
    LV_Shadow.Height = LV_Control.Height + 2 * LV_Width
    
End Sub

Public Sub Form_Resize_Control(LV_Form As Form, LV_Control As Control)

Dim lvWidth             As Long

    LV_Control.Width = LV_Form.ScaleWidth - 2 * LV_Control.Left
    
End Sub

Public Sub Form_Move_Controls_To_Center(LV_Form As Form, Optional ControlType As String)

Dim LV_Control              As Control
Dim lvLeft                  As Long
Dim lvTop                   As Long
Dim lvRight                 As Long
Dim lvBottom                As Long
Dim lvStart                 As Boolean
Dim lvWidth                 As Long
Dim lvHeight                As Long
Dim lvOffsetHori            As Long
Dim lvOffsetVert            As Long

    lvStart = False
    For Each LV_Control In LV_Form.Controls
        If ControlHasPositionProperty(LV_Control) = True And _
                ControlTypeFit(LV_Control, ControlType) = True Then
            If lvStart = False Then
                lvStart = True
                lvLeft = LV_Control.Left
                lvTop = LV_Control.Top
                lvRight = LV_Control.Left + LV_Control.Width
                lvBottom = LV_Control.Top + LV_Control.Height
            Else
                If lvLeft > LV_Control.Left Then
                    lvLeft = LV_Control.Left
                End If
                If lvTop > LV_Control.Top Then
                    lvTop = LV_Control.Top
                End If
                If lvRight < LV_Control.Left + LV_Control.Width Then
                    lvRight = LV_Control.Left + LV_Control.Width
                End If
                If lvBottom < LV_Control.Top + LV_Control.Height Then
                    lvBottom = LV_Control.Top + LV_Control.Height
                End If
            End If
        End If
    Next
    lvWidth = lvRight - lvLeft
    lvHeight = lvBottom - lvTop
    lvOffsetHori = (LV_Form.ScaleWidth - lvWidth) / 2 - lvLeft
    lvOffsetVert = (LV_Form.ScaleHeight - lvHeight) / 2 - lvTop
    
    Form_Move_Controls LV_Form, lvOffsetHori, lvOffsetVert, ControlType
    
End Sub

Public Sub Form_Move_Controls(LV_Form As Form, lvOffsetHori As Long, _
                                lvOffsetVert As Long, _
                                Optional ControlType As String)
                                
Dim LV_Control              As Control

    For Each LV_Control In LV_Form.Controls
        If ControlHasPositionProperty(LV_Control) = True And _
                ControlTypeFit(LV_Control, ControlType) = True Then
            LV_Control.Left = LV_Control.Left + lvOffsetHori
            LV_Control.Top = LV_Control.Top + lvOffsetVert
        End If
    Next
    
End Sub



