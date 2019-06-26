Attribute VB_Name = "modSpecialFolders"
'---------------------------------------------------------------------------------------
' Module    : modSpecialFolders
' Author    : Leo Herrera
' Date      : 29/10/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

'Form Code
Public Const CSIDL_DESKTOP = &H0 '// The Desktop - virtual folder
Public Const CSIDL_PROGRAMS = 2 '// Program Files
Public Const CSIDL_CONTROLS = 3 '// Control Panel - virtual folder
Public Const CSIDL_PRINTERS = 4 '// Printers - virtual folder
Public Const CSIDL_DOCUMENTS = 5 '// My Documents
Public Const CSIDL_FAVORITES = 6 '// Favourites
Public Const CSIDL_STARTUP = 7 '// Startup Folder
Public Const CSIDL_RECENT = 8 '// Recent Documents
Public Const CSIDL_SENDTO = 9 '// Send To Folder
Public Const CSIDL_BITBUCKET = 10 '// Recycle Bin - virtual folder
Public Const CSIDL_STARTMENU = 11 '// Start Menu
Public Const CSIDL_DESKTOPFOLDER = 16 '// Desktop folder
Public Const CSIDL_DRIVES = 17 '// My Computer - virtual folder
Public Const CSIDL_NETWORK = 18 '// Network Neighbourhood - virtual folder
Public Const CSIDL_NETHOOD = 19 '// NetHood Folder
Public Const CSIDL_FONTS = 20 '// Fonts folder
Public Const CSIDL_SHELLNEW = 21 '// ShellNew folder


'Module Code
Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, ByVal pszPath As String) As Long

Public Type SHITEMID
    cb As Long
    abID As Byte
End Type

Public Type ITEMIDLIST
    mkid As SHITEMID
End Type

Public Const MAX_PATH As Integer = 260

Public Function fGetSpecialFolder(CSIDL As Long, hWnd As Long) As String

Dim sPath As String
Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
    fGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(hWnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
        End If
    End If

End Function

'Private Sub form_load()
'    MsgBox "Desktop Folder " & fGetSpecialFolder(CSIDL_DESKTOPFOLDER)
'    MsgBox "Recent Folder " & fGetSpecialFolder(CSIDL_RECENT)
'    '// etc...
'End Sub

