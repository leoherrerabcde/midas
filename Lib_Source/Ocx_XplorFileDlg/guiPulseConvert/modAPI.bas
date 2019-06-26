Attribute VB_Name = "modAPI"
'-----------
'Author: Sathyaish Chakravarthy
'Date: 5th February, 2002.
'Contact Information:
'                       ''''''''''''''''''''''''''''''''''''''''''''
'                       'Email: SathyaishC@yahoo.co.uk             '
'                       '       VisualBasicLearner@ yahoo.com      '
'                       'Web:   SathyaishC.tripod.com              '
'                       ''''''''''''''''''''''''''''''''''''''''''''
'Purpose of the Project: The object of this sample project is to demonstrate how
'to use the EnumWindows API function with an application-defined CallBack function
'in order to enumerate all the top-level windows running in the Windows environment
'at any point of time.
'The enumeration of windows is done using the EnumWindows API function and an application
'defined callback function called fnEnumWindowsCallback. The EnumWindow API function returns
'searches for a top-level window and then passes the handle of the window found to the application
'defined function called fnEnumWindowsCallback, which in turn processes the window handle value
'recieved to gather information about the 'Title/Caption' of the window recieved from EnumWindows
'function.
'The sample also demonstrates how you can shut down windows from within your Visual Basic
'Application. It provides a form interface somewhat similar to the Windows Task Manager.
'It also provides you the flexibility to kill an application from within your VB applicaiton.
'However, the list of running top-level windows is not immediately updated after you kill
'an application.

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function EnumWindows Lib "user32" _
(ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_QUIT = &H12
Public Const WM_DESTROY = &H2

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_SHUTDOWN = 1

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long


Public ArrWindowList() As String
Public LngNumWindows As Long


Public Function fnEnumWindowsCallback(ByVal hwnd As Long, ByVal lParam As Long) As Long

'This function is the application defined callback function that recieves a window
'handle of a top-level window each time from the EnumWindows API, it then processes
'the window handle in a manner it decides (in this case, by retrieving the title/caption
'to the window by calling the GetWindowText API) and then returns control to Windows/
'inside the EnumWindows function which again passes it another handle. This process
'continues until no more windows are left and that is when the callback function
'fnEnumWindowsCallback returns False in order to stop enumeration.

Dim LngCharLength As Long
Dim StrWindowText As String

If hwnd = 0 Then
    fnEnumWindowsCallback = False
    Exit Function
End If
StrWindowText = String(255, vbNullChar)
LngCharLength = GetWindowText(hwnd, StrWindowText, 255)
StrWindowText = Left(StrWindowText, LngCharLength)


If Trim(StrWindowText) <> vbNullString Then
    If IsWindowVisible(hwnd) Then
        ReDim Preserve ArrWindowList(LngNumWindows)
        ArrWindowList(LngNumWindows) = StrWindowText
        LngNumWindows = LngNumWindows + 1
    End If
End If

fnEnumWindowsCallback = True

End Function
