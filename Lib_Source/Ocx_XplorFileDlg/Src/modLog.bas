Attribute VB_Name = "modLog"
Option Explicit

Public GV_hFile     As Integer
Public GV_Path_Iconos   As String
Public GV_File_Btn_Nsuperior    As String

Sub OpenLogFile()

Dim lv_LogFile      As String
Dim lvPath          As String

   'On Error GoTo OpenLogFile_Error

'On Error GoTo Exit_OpenLogFile

    If GetSettingBooleanParameter(GC_ENABLE_DBG_FILE_LOG, False) = False Then
        If GV_hFile <> 0 Then
            Close #GV_hFile
            GV_hFile = 0
        End If
        Exit Sub
    End If
    If GV_hFile = 0 Then
        GV_hFile = FreeFile
        lv_LogFile = "OpenDlg_" + Format(Now(), "YYYY_MM_SS_hh_mm_ss") + ".log"
        lvPath = Retroceder_Path(App.Path)
        lvPath = lvPath & "\Dbg"
          
        Open lvPath & "\" & lv_LogFile For Output As GV_hFile
        Print #GV_hFile, "App.Title : " & App.Title
        Print #GV_hFile, "App.Path : " & App.Path
        Print #GV_hFile, "Time Start: " & Format(Now(), "hh:mm:ss")
        Print #GV_hFile, vbCrLf
    End If
Exit_OpenLogFile:

   'On Error GoTo 0
   Exit Sub

OpenLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure OpenLogFile of Módulo modLog"

End Sub

Sub WriteLogFile(lsStr As String)

   'On Error GoTo WriteLogFile_Error

    OpenLogFile
    If GV_hFile Then
        Print #GV_hFile, Format(Now, "yyyy-mm-dd hh:mm:ss : ") & lsStr
    End If

   'On Error GoTo 0
   Exit Sub

WriteLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure WriteLogFile of Módulo modLog"

End Sub


Sub CloseLogFile()

'On Error GoTo Exit_CloseLogFile

    If GV_hFile Then
        Print #GV_hFile, vbCrLf
        Print #GV_hFile, "App.Title :" & App.Title
        Print #GV_hFile, "App.Path :" & App.Path
        Print #GV_hFile, "Time End: " & Format(Now(), "hh:mm:ss")
        Close GV_hFile
    End If
    
Exit_CloseLogFile:
        GV_hFile = 0

End Sub
