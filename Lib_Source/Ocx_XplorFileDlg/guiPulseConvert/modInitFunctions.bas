Attribute VB_Name = "modInitFunctions"
'---------------------------------------------------------------------------------------
' Module    : modInitFunctions
' Author    : lherrera
' Date      : 23/02/2011
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Sub Init_Vars()

    On Error Resume Next
    If GetSettingBooleanParameter(GC_ERASE_DEBUG_SETTING, False) = True Then
        DeleteSetting App.Title, GC_DEBUG
        SetSettingBooleanParameter GC_ERASE_DEBUG_SETTING, True
    End If
    'On Error GoTo 0
    
    InitTimeOutCounters
    
    Set m_Project = New ClassProject
    
    Set GV_clsErrorView = New clsInstanceControl
    Set GV_clsSecondInstance = New clsInstanceControl
    Set GV_clsTemplateConfigSpreadSheet = New clsInstanceControl
    Set GV_clsOpenPulseDialog = New clsInstanceControl
    Set GV_clsSpreadSheetView = New clsInstanceControl
    Set GV_clsExportSpreadSheet = New clsInstanceControl
    Set GV_clsConfigSpreadSheet = New clsInstanceControl
    Set GV_clsStart = New clsInstanceControl
    Set GV_clsStartUp = New clsInstanceControl
    Set GV_clsProjectSelLocation = New clsInstanceControl
    Set GV_clsPjtFromFile = New clsInstanceControl
    Set GV_clsPjtFromList = New clsInstanceControl
    Set m_MakeRound = New clsMakeRound
    
    BackGroundProjectList.Count = 0
    BackGroundProjectList.ListIndex = -1
    
    GV_Project_Opened = False
    GV_Project_Closed = True
    
    m_MakeRound.SetRoundValue 20
    
    'GV_Config_Path = App.Path & "\" & GC_CONFIG_PATH
    GV_Config_Path = Set_Config_Path(Retroceder_Path(App.Path))
    GV_XlsDll_Path = Set_XlsDll_Path(Retroceder_Path(App.Path))
    GV_Bin2Xls_App = "cvtBin2xls.exe"
    GV_Bin2Xls_App = GetSetting(App.Title, _
                                GC_CONFIGURATION_SECTION, _
                                GC_BIN_2_XLS_APP, _
                                GV_Bin2Xls_App)
    SaveSetting App.Title, _
                GC_CONFIGURATION_SECTION, _
                GC_BIN_2_XLS_APP, _
                GV_Bin2Xls_App
    
    If GetSettingBooleanParameter(GC_ENABLE_DBG_DLL_FILE_LOG, False) = True Then
        Pulse_Log_Path GV_XlsDll_Path
    End If
    
    GV_ListProjectFile = GV_Config_Path & "\list_project.lpp"
    
    GV_PrevInstance = Is_PrevInstance
    
    GV_Debug.Template = True
    
    GV_Index_Pjt = 0
    
    Init_Message_Length

    Debug_Save_Configuration
    
End Sub

Sub Debug_Save_Configuration()

    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "App.Path", App.Path
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "GV_Config_Path", GV_Config_Path
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "GV_ListProjectFile", GV_ListProjectFile
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "GV_PrevInstance", GV_PrevInstance

End Sub

Sub Init_Message_Length()

Dim lvUbound            As Long

    lvUbound = Message_Header_Const.MSG_END_PROJECT - _
                Message_Header_Const.MSG_RUNXLS
    
    ReDim GV_Msg_Len(lvUbound)
    ReDim GV_Msg_Header(lvUbound)
    
    GV_Msg_Len(Message_Header_Const.MSG_RUNXLS) = 4
    GV_Msg_Len(Message_Header_Const.MSG_ALIVE) = 1
    GV_Msg_Len(Message_Header_Const.MSG_START_PROJECT) = 1
    GV_Msg_Len(Message_Header_Const.MSG_FILE_START) = 3
    GV_Msg_Len(Message_Header_Const.MSG_STATUS) = 3
    GV_Msg_Len(Message_Header_Const.MSG_SAVING_FILE) = 1
    GV_Msg_Len(Message_Header_Const.MSG_XLS_FILE_READY) = 1
    GV_Msg_Len(Message_Header_Const.MSG_END_PROJECT) = 0

    GV_Msg_Header(Message_Header_Const.MSG_RUNXLS) = "RUNXLS"
    GV_Msg_Header(Message_Header_Const.MSG_START_PROJECT) = "START_PROJECT"
    GV_Msg_Header(Message_Header_Const.MSG_FILE_START) = "FILE_START"
    GV_Msg_Header(Message_Header_Const.MSG_STATUS) = "STATUS"
    GV_Msg_Header(Message_Header_Const.MSG_SAVING_FILE) = GC_MSG_SAVING_FILE
    GV_Msg_Header(Message_Header_Const.MSG_XLS_FILE_READY) = "XLS_FILE_READY"
    GV_Msg_Header(Message_Header_Const.MSG_END_PROJECT) = "END_PROJECT"
    
End Sub

Function Set_Flt_Dts_Path(Optional LV_Path As String) As String

Dim lv_Path_Setting             As String
Dim lv_New_Path                 As String

    If LV_Path <> "" Then
        lv_Path_Setting = LV_Path
    Else
        lv_Path_Setting = App.Path
    End If
    
    
    On Error Resume Next
    Do
        Set_Flt_Dts_Path = lv_Path_Setting & "\Templates\FltDtsTemplates"
        If Is_Folder(Set_Flt_Dts_Path) = True Then
            Exit Function
        Else
            lv_Path_Setting = Retroceder_Path(lv_Path_Setting)
            If lv_Path_Setting = "" Then
                Set_Flt_Dts_Path = ""
                Exit Function
            End If
        End If
    Loop

End Function

Function Set_XlsDll_Path(Optional LV_Path As String) As String

Dim lv_Path_Setting             As String
Dim lv_New_Path                 As String

    Set_XlsDll_Path = ""
    If GetSettingBooleanParameter("Debug GC_DLL_PATH", False) = True Then
        Set_XlsDll_Path = GetSetting(App.Title, _
                                    GC_CONFIGURATION_SECTION, _
                                    GC_DLL_PATH, _
                                    LV_Path & "\" & GC_DLL_PATH)
        SaveSetting App.Title, GC_CONFIGURATION_SECTION, GC_DLL_PATH, Set_XlsDll_Path
        Exit Function
    End If
    
    If LV_Path <> "" Then
        lv_Path_Setting = LV_Path
    Else
        lv_Path_Setting = App.Path
    End If
    
    Set_XlsDll_Path = ""
    
    On Error Resume Next
    Do
        Set_XlsDll_Path = lv_Path_Setting & "\" & GC_DLL_PATH
        If Is_Folder(Set_XlsDll_Path) = True Then
            Exit Do
        Else
            lv_Path_Setting = Retroceder_Path(lv_Path_Setting)
            If lv_Path_Setting = "" Then
                Set_XlsDll_Path = ""
                Exit Function
            End If
        End If
    Loop
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, GC_DLL_PATH, Set_XlsDll_Path
    
End Function

Function Set_Config_Path(Optional LV_Path As String) As String

Dim lv_Path_Setting             As String
Dim lv_New_Path                 As String

    If LV_Path <> "" Then
        lv_Path_Setting = LV_Path
    Else
        lv_Path_Setting = App.Path
    End If
    
    
    On Error Resume Next
    Do
        Set_Config_Path = lv_Path_Setting & "\" & GC_CONFIG_PATH
        SaveSetting App.Title, GC_CONFIGURATION_SECTION, "Config_Path", Set_Config_Path
        If Is_Folder(Set_Config_Path) = True Then
            Exit Function
        Else
            lv_Path_Setting = Retroceder_Path(lv_Path_Setting)
            If lv_Path_Setting = "" Then
                Set_Config_Path = ""
                Exit Function
            End If
        End If
    Loop
    
End Function

Function Find_Folder(lv_Folder As String, Optional LV_Path As String) As String

Dim lv_Path_Setting             As String
Dim lv_New_Path                 As String

    If LV_Path <> "" Then
        lv_Path_Setting = LV_Path
    Else
        lv_Path_Setting = App.Path
    End If
    
    
    On Error Resume Next
    Do
        Find_Folder = lv_Path_Setting & "\" & lv_Folder
        If Is_Folder(Find_Folder) = True Then
            Exit Function
        Else
            lv_Path_Setting = Retroceder_Path(lv_Path_Setting)
            If lv_Path_Setting = "" Then
                Find_Folder = ""
                Exit Function
            End If
        End If
    Loop

End Function

Function Set_Path_Settings(Optional LV_Path As String) As String

Dim lv_Path_Setting             As String
Dim lv_New_Path                 As String

    If LV_Path <> "" Then
        lv_Path_Setting = LV_Path
    Else
        lv_Path_Setting = App.Path
    End If
    
    
    'On Error Resume Next
    Do
        Set_Path_Settings = lv_Path_Setting & "\Settings"   '\Settings.txt"
        'If GetAttr(Set_Path_Settings) And vbArchive Then
        If Is_Folder(Set_Path_Settings) = True Then
            Exit Function
        Else
            lv_Path_Setting = Retroceder_Path(lv_Path_Setting)
            If lv_Path_Setting = "" Then
                Set_Path_Settings = ""
                Exit Function
            End If
        End If
    Loop

End Function

Function Set_File_Settings(Optional LV_Path As String) As String

Dim lv_Path_Setting             As String
Dim lv_New_Path                 As String

    If LV_Path <> "" Then
        lv_Path_Setting = LV_Path
    Else
        lv_Path_Setting = App.Path
    End If
    
    
    'On Error Resume Next
    Do
        Set_File_Settings = lv_Path_Setting & "\Settings"   '\Settings.txt"
        'If GetAttr(Set_File_Settings) And vbArchive Then
        If Is_Folder(Set_File_Settings) = True Then
            Set_File_Settings = Set_File_Settings & "\Settings.txt"
            Exit Function
        Else
            lv_Path_Setting = Retroceder_Path(lv_Path_Setting)
            If lv_Path_Setting = "" Then
                Set_File_Settings = ""
                Exit Function
            End If
        End If
    Loop
        
End Function

Function Set_Form_File_Settings(LV_Name As String) As String

Dim lv_Path_Setting             As String
Dim lv_New_Path                 As String

    lv_Path_Setting = App.Path
    On Error Resume Next
    Do
        Set_Form_File_Settings = lv_Path_Setting & "\Settings\" & LV_Name & "_Settings.txt"
        If GetAttr(Set_Form_File_Settings) And vbArchive Then
            Exit Function
        Else
            lv_Path_Setting = Retroceder_Path(lv_Path_Setting)
            If lv_Path_Setting = "" Then
                Set_Form_File_Settings = ""
                Exit Function
            End If
        End If
    Loop
    
    
End Function


Sub InitTimeOutCounters()

    GV_TimeToKillSecond = 5000 * (1 + Rnd)
    GV_TimeToKillMySelf = 5000 * (1 + Rnd)
    GV_TimeOutComm = 5000 * (1 + Rnd)
    GV_TimeOutSendAlive = 2000 * (1 + Rnd)
    
End Sub
