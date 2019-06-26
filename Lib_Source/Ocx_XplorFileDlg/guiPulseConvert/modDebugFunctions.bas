Attribute VB_Name = "modDebugFuntions"
'---------------------------------------------------------------------------------------
' Module    : modDebugFuntions
' Author    : Leo Herrera
' Date      : 16/05/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_hFile            As Integer

Sub SaveSettingDebug(lvKey As String, lsValue)

    SaveSetting App.Title, GC_DEBUG, GetTickCount & "_" & lvKey, lsValue

End Sub


'Sub Write_Log(lsData As String, Optional lsFunction As String = "")
'
'    If GetSettingBooleanParameter(GC_ENABLE_EXPORTFILE_LOG, False) = True Then
'        If PV_hFile = 0 Then
'            If GV_PrevInstance = False Then
'                lsName = "Dbg_Main_" & Format(Now(), "hh_mm_ss") & ".log"
'            Else
'                lsName = "Dbg_Scnd_" & Format(Now(), "hh_mm_ss") & ".log"
'            End If
'            lsName = Retroceder_Path(App.Path) & "\Exe\" & lsName
'            PV_hFile = FreeFile
'            Open lsName For Append As PV_hFile
'        End If
'        If lsFunction <> "" Then
'            lsStr = Format(Now(), "hh:mm:ss ") & "| Fn " & lsFunction & "->" & lsData
'        Else
'            lsStr = Format(Now(), "hh:mm:ss -> ") & lsData
'        End If
'        Print #PV_hFile, lsStr
'    End If
'
'End Sub
'
'Sub Close_Log()
'
'    If PV_hFile Then
'        Close #PV_hFile
'    End If
'
'End Sub
