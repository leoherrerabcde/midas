Attribute VB_Name = "modMain"
Option Explicit

Sub Main()

    Set GV_Mdi = New MDIMain
    
    'app.Revision
    'SaveSettingDebug "App.hInstance", App.hInstance
    
    WriteLogFile "Loading frmMain"
    Load GV_Mdi

    WriteLogFile "Verify Previous Instance"
    If GV_PrevInstance = False Then
        WriteLogFile "No Previous Instance"
        Forzar_PrevInstance
        GV_Mdi.Show
    Else
        WriteLogFile "There is another Instance"
        If GetSettingBooleanParameter(GC_VISIBLE_SND_INSTANCE, False) = True Then
            GV_Mdi.Show
        End If
        GV_Send_Progress = True
    End If
    
    WriteLogFile "frmMain Loaded"
    
End Sub

Sub Forzar_PrevInstance()

Dim lv_Counter          As Integer

    lv_Counter = GetSetting(App.Title, GC_CONFIGURATION_SECTION, GV_Mdi.Name & ".app.PrevInstance", 0)
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, GV_Mdi.Name & ".app.PrevInstance", lv_Counter + 1
    
End Sub

Function Is_PrevInstance() As Boolean

    Is_PrevInstance = App.PrevInstance
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "App.PrevInstance", App.PrevInstance
    If Is_PrevInstance = False Then
        If Find_PrevInstance(App.Title) = True Then
            Is_PrevInstance = True
        End If
    End If
    
End Function

Sub Release_PrevInstance()

Dim lv_Counter          As Integer

    lv_Counter = GetSetting(App.Title, GC_CONFIGURATION_SECTION, GV_Mdi.Name & ".app.PrevInstance", 0)
    If lv_Counter Then
        SaveSetting App.Title, GC_CONFIGURATION_SECTION, GV_Mdi.Name & ".app.PrevInstance", lv_Counter - 1
    End If
    
End Sub


