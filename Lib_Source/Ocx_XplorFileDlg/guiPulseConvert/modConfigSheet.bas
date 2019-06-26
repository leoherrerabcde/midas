Attribute VB_Name = "modConfigSheet"
'---------------------------------------------------------------------------------------
' Module    : modConfigSheet
' Author    : Leo Herrera
' Date      : 27/11/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Sub Cpy_Column_Config(Lv_Dest As typeConfigSheetColumns, Lv_Src As typeConfigSheetColumns)

Dim i           As Integer

    With Lv_Src
        Lv_Dest.Count = .Count
        ReDim Lv_Dest.Column(.Count - 1)
        Lv_Dest.ColumnConfigName = .ColumnConfigName
        For i = 0 To .Count - 1
            Lv_Dest.Column(i).ColumnName = .Column(i).ColumnName
            Lv_Dest.Column(i).Order = .Column(i).Order
            Lv_Dest.Column(i).Visible = .Column(i).Visible
        Next
    End With

End Sub

Sub Cpy_Spread_Config(Lv_Dest As typeConfigSpreadSheet, _
                        Lv_Src As typeConfigSpreadSheet)
    
    Lv_Dest.byFiles = Lv_Src.byFiles
    Lv_Dest.byInterval = Lv_Src.byInterval
    Lv_Dest.byPulses = Lv_Src.byPulses
    Lv_Dest.IntervalPerSheet = Lv_Src.IntervalPerSheet
    Lv_Dest.PulsesPerSheet = Lv_Src.PulsesPerSheet
    Lv_Dest.SheetsPerSpreadSheet = Lv_Src.SheetsPerSpreadSheet
    Lv_Dest.SpreadConfigName = Lv_Src.SpreadConfigName
    
End Sub

Sub SetDefault(optFiles As OptionButton, optPulses As OptionButton, _
                optInterval As OptionButton, txtSheetsPerSpread As TextBox, _
                txtPulsesPerSheet As TextBox, txtIntervalPerSheet As TextBox)

End Sub
Function Get_Index_CboBx(lsTxt As String, Cbo As ComboBox) As Integer

Dim i           As Integer

    For i = 0 To Cbo.ListCount - 1
        If Cbo.List(i) = lsTxt Then
            Get_Index_CboBx = i
            Exit Function
        End If
    Next
    Get_Index_CboBx = -1
    
End Function

Sub Load_Config_Column(ConfigList As typeColumnsConfigList, _
                            Cbo As ComboBox)

Dim LV_ConfigFileName           As String
Dim h                           As Integer
Dim i                           As Integer

    'On Error GoTo Load_Config_Column_Error

    ConfigList.Count = 0
    LV_ConfigFileName = GV_Config_Path
    If Is_Folder(LV_ConfigFileName) = True Then
        LV_ConfigFileName = LV_ConfigFileName & "\ConfigColumnList.cfg"
        h = FreeFile
        Open LV_ConfigFileName For Binary Access Read As h
        Get #h, , ConfigList
        Close #h
        For i = 0 To ConfigList.Count - 1
            Cbo.AddItem ConfigList.Config(i).ColumnConfigName
        Next
    End If
    
    'On Error GoTo 0
    Exit Sub

Load_Config_Column_Error:

End Sub

Function CalcPulsesToConfig(lsPulses As String) As String

Dim lPulses             As Long
Dim K_Pulses            As Long
Dim U_Pulses            As Long

    If IsNumeric(lsPulses) = True Then
        lPulses = lsPulses
        K_Pulses = lPulses / 1000
        U_Pulses = lPulses - 1000 * K_Pulses
        If K_Pulses And U_Pulses = 0 Then
            CalcPulsesToConfig = Trim$(Str(K_Pulses)) & "K"
        Else
            CalcPulsesToConfig = Trim$(Str(lPulses))
        End If
    Else
        CalcPulsesToConfig = ""
    End If
End Function

Function CalcIntervalToConfig(lsInterval) As String

Dim lMin                As Long
Dim lSec                As Long
Dim l_msec              As Long
Dim dInterval           As Double

    If IsNumeric(lsInterval) = True Then
        dInterval = lsInterval
        If dInterval >= 60 Then
            lMin = Int(dInterval / 60)
            CalcIntervalToConfig = Trim$(Str(lMin)) & "min"
            dInterval = dInterval - lMin * 60
        Else
            CalcIntervalToConfig = ""
        End If
        lSec = Fix(dInterval)
        If lSec Then
            CalcIntervalToConfig = CalcIntervalToConfig & _
                                    Trim$(Str(lSec)) & "sec"
        End If
'        l_msec = 1000 * (dInterval - lSec)
'        If l_msec Then
'            CalcIntervalToConfig = CalcIntervalToConfig & _
'                                    Trim$(Str(l_msec)) & "ms"
'        End If
    Else
        CalcIntervalToConfig = ""
    End If
    
End Function

Function CalcLPulsesToConfig(ByVal lPulses As Long) As String

Dim K_Pulses            As Long
Dim U_Pulses            As Long

    K_Pulses = lPulses / 1000
    U_Pulses = lPulses - 1000 * K_Pulses
    If K_Pulses And U_Pulses = 0 Then
        CalcLPulsesToConfig = Trim$(Str(K_Pulses)) & "K"
    Else
        CalcLPulsesToConfig = Trim$(Str(lPulses))
    End If

End Function

Function CalcDIntervalToConfig(ByVal dInterval As Long) As String

Dim lMin                As Long
Dim lSec                As Long
Dim l_msec              As Long

    If dInterval >= 60 Then
        lMin = Int(dInterval / 60)
        CalcDIntervalToConfig = Trim$(Str(lMin)) & "min"
        dInterval = dInterval - lMin * 60
        lSec = Fix(dInterval)
    Else
        CalcDIntervalToConfig = ""
        lSec = dInterval
    End If
    If lSec Then
        CalcDIntervalToConfig = CalcDIntervalToConfig & _
                                Trim$(Str(lSec)) & "sec"
    End If
    
End Function

Sub SetSpreadConfigName(SpreadConfig As typeConfigSpreadSheet)

Dim lvName          As String
Dim i               As Integer

    With SpreadConfig
        If .byFiles = True Then
                lvName = "A"
        Else
            If .byPulses = True Then
                lvName = "P" & CalcLPulsesToConfig(.PulsesPerSheet)
            Else
                If .byInterval = True Then
                    lvName = "T" & CalcDIntervalToConfig(.IntervalPerSheet)
                Else
                    lvName = ""
                    .SpreadConfigName = lvName
                    Exit Sub
                End If
            End If
        End If
        lvName = lvName & "_H" & Trim$(Str(.SheetsPerSpreadSheet))
        .SpreadConfigName = lvName
    End With
    
End Sub


Sub CalcConfigNameOnList(ConfigList As typeConfigSpreadSheetList)

Dim i               As Integer

    For i = 1 To ConfigList.Count - 1
        SetSpreadConfigName ConfigList.ConfigList(i)
    Next
    
End Sub

Sub Load_ConfigSpreadSheet(ConfigList As typeConfigSpreadSheetList, _
                            Cbo As ComboBox)

Dim LV_ConfigFileName           As String
Dim h                           As Integer
Dim i                           As Integer

    ConfigList.Count = 0
    'On Error GoTo Error_ConfigSpreadFile_Doesnt_Exist
    LV_ConfigFileName = GV_Config_Path
    If Is_Folder(LV_ConfigFileName) = True Then
        LV_ConfigFileName = LV_ConfigFileName & "\ConfigSheetsList.cfg"
        h = FreeFile
        Open LV_ConfigFileName For Binary Access Read As h
        Get #h, , ConfigList
        Close #h
        CalcConfigNameOnList ConfigList
        For i = 0 To ConfigList.Count - 1
            Cbo.AddItem ConfigList.ConfigList(i).SpreadConfigName
        Next
    End If
    Exit Sub
    
Error_ConfigSpreadFile_Doesnt_Exist:
    'On Error GoTo 0
    
End Sub

Sub Save_ConfigSpreadSheet(ConfigList As typeConfigSpreadSheetList)

Dim LV_ConfigFileName           As String
Dim h                           As Integer

    LV_ConfigFileName = GV_Config_Path
    If Is_Folder(LV_ConfigFileName) = False Then
        Create_Folder LV_ConfigFileName
    End If
    LV_ConfigFileName = LV_ConfigFileName & "\ConfigSheetsList.cfg"
    h = FreeFile
    Open LV_ConfigFileName For Binary Access Write As h
    Put #h, , ConfigList
    Close #h

End Sub

Sub Save_Config_Column(ConfigList As typeColumnsConfigList)

Dim LV_ConfigFileName           As String
Dim h                           As Integer

    LV_ConfigFileName = GV_Config_Path
    If Is_Folder(LV_ConfigFileName) = False Then
        Create_Folder LV_ConfigFileName
    End If
    LV_ConfigFileName = LV_ConfigFileName & "\ConfigColumnList.cfg"
    h = FreeFile
    Open LV_ConfigFileName For Binary Access Write As h
    Put #h, , ConfigList
    Close #h

End Sub


Function GetIndexSpreadConfig(Config As typeConfigSpreadSheet, _
                        ConfigList As typeConfigSpreadSheetList) As Integer

Dim i           As Integer

    GetIndexSpreadConfig = -1
    For i = 0 To ConfigList.Count - 1
        If ConfigList.ConfigList(i).SpreadConfigName = Config.SpreadConfigName Then
            GetIndexSpreadConfig = i
        End If
    Next
    
End Function

Function GetIndexColumnConfig(Config As typeConfigSheetColumns, _
                        ConfigList As typeColumnsConfigList) As Integer

Dim i           As Integer

    GetIndexColumnConfig = -1
    For i = 0 To ConfigList.Count - 1
        If ConfigList.Config(i).ColumnConfigName = Config.ColumnConfigName Then
            GetIndexColumnConfig = i
        End If
    Next
    
End Function


