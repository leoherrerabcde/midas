Attribute VB_Name = "modLoadRes"
'---------------------------------------------------------------------------------------
' Module    : modLoadRes
' Author    : Leo Herrera
' Date      : 03/01/2014
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Sub LoadResControls(frm As Controls)

    On Error Resume Next

    Dim obj As Control

    For Each obj In frm
        If obj.Tag <> "" Then
            If IsNumeric(obj.Tag) = True Then
                obj.Caption = LoadResString(CInt(obj.Tag))
            End If
        End If
    Next

End Sub

Sub LoadResStrings(frm As Form)
    
    On Error Resume Next

    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer

    'set the form's caption
    frm.Caption = LoadResString(CInt(frm.Tag))
    

    'set the font
    Set fnt = frm.Font
    fnt.Name = LoadResString(20)
    fnt.Size = CInt(LoadResString(21))
    

    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In frm.Controls
        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = Val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next


End Sub

