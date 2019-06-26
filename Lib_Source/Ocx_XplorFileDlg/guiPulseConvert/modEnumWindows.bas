Attribute VB_Name = "modEnumWindows"
'---------------------------------------------------------------------------------------
' Module    : modEnumWindows
' Author    : Leo Herrera
' Date      : 21/12/2012
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public Function Find_App(lsApp As String) As Boolean

Dim i               As Integer
Dim Count           As Integer
Dim Static_Count    As Integer
Dim v               As Integer

    RefreshList
    Static_Count = Static_Count + 1
    
    Count = 0
    For i = 0 To UBound(ArrWindowList)
        v = InStr(ArrWindowList(i), lsApp)
        If v Then
            Count = Count + 1
        End If
    Next
    
    Find_App = False
    If Count >= 1 Then
        Find_App = True
    End If
    
End Function

Public Function Find_PrevInstance(lsApp As String) As Boolean

Dim i           As Integer
Dim Count       As Integer
Dim Static_Count    As Integer

    RefreshList
    Static_Count = Static_Count + 1
    WriteLogFile "Find_PrevInstance: " & lsApp
    Count = 0
    'SaveSettingDebug "ArrWindowList", Static_Count
    For i = 0 To UBound(ArrWindowList)
        'SaveSettingDebug "ArrWindowList(" & Trim$(i) & ")", ArrWindowList(i)
        'If Left$(ArrWindowList(i), Len(lsApp)) = lsApp Then
        If ArrWindowList(i) = lsApp Then
            'SaveSettingDebug lsApp & "=" & "ArrWindowList(" & Trim$(i) & ")", ArrWindowList(i)
            WriteLogFile "App " & ArrWindowList(i) & " is running"
            
            Count = Count + 1
        End If
    Next
    
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "Find_PrevInstance", Count
    SaveSetting App.Title, GC_CONFIGURATION_SECTION, "Finding_App", lsApp
    'SaveSettingDebug "ArrWindowList_Fin", Count
    
    Find_PrevInstance = False
    If Count > 1 Then
        Find_PrevInstance = True
    End If
    
    If GetSettingBooleanParameter(GC_INVERSE_PREV_INSTANCE, False) = True Then
        If Find_PrevInstance = True Then
            Find_PrevInstance = False
        Else
            Find_PrevInstance = True
        End If
    End If
    
End Function

Public Sub RefreshList()

Dim LngRetVal As Long
Dim LngCounter As Long
Dim LngNumWindows As Long

'Set colWindowList = New Collection

Erase ArrWindowList
modAPI.LngNumWindows = 0


LngRetVal = EnumWindows(AddressOf fnEnumWindowsCallback, 0&)

End Sub


Public Sub KillApp(lsApp As String)

Dim hWindow As Long
Dim LngRetVal As Long
    
    hWindow = FindWindow(vbNullString, lsApp)
    'Debug.Assert hWindow <> 0
    If hWindow <> 0 Then
        LngRetVal = PostMessage(hWindow, WM_QUIT, 0, 0)
    'Debug.Assert LngRetVal <> 0
    'Call RefreshList
    End If

End Sub
