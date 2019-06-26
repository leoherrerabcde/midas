VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OcxDirCboImg 
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   ScaleHeight     =   2580
   ScaleWidth      =   3990
   Begin MSComctlLib.ListView LstVwFolders 
      Height          =   1575
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2778
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.DriveListBox DriveDir 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.DirListBox DirList 
      Height          =   1440
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.FileListBox FileList 
      Height          =   1260
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageCombo ImageComboDir 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageComboDir"
      ImageList       =   "ilsIcons16"
   End
   Begin MSComctlLib.ImageList ilsIcons16 
      Left            =   3120
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":0000
            Key             =   "MyComputer"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":08A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":0CFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsIcons32 
      Left            =   2520
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":34B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1920
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":3904
            Key             =   "Up One Level"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":3A16
            Key             =   "Clsdfold"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":3D30
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":3E42
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":3F54
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":4066
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OcxDirCboImg.ctx":4178
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OcxDirCboImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : OcxDirCboImg
' Author    : Leo Herrera
' Date      : 18/10/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_Last_Path            As String
Private PV_Selected_Path        As String

Private fso As New FileSystemObject
Private d, dc, s, n, t
Private fs, f, f1, fc, s1, nf

Private bIsDrive As Boolean
Private ci As ComboItem
Private Children() As String
Private ChildIndex As Variant

Private imgX As ListImage
Private itmFldr As ListItem
Private itmX As ListItem
Private iListCount As Integer
Private iParentIndex As Integer
Private iNumOfChildren          As Integer
Private sCheckChild             As String
Private sDir                    As String
Private sItem                   As String
Private sKey                    As String
Private sKey1                   As String
Private sPath                   As String
Private sText                   As String
Private PV_VirtualFolderList()  As Long
'Property Let Options(lvOption As DirCboFileConstants)
'
'End Property

Public Event ItemSelected(lvPath As String)
Public Event VirtualFolderSelected(CSIDL As Long)
Public Event ListClosed()
Public Event SelectionChanged()

Property Get Version() As String

    Version = App.Title & ":" & App.Major & "," & App.Minor & "," & App.Revision
    
End Property

Property Get SelectedPath() As String

    With UserControl
        If IsPathDrive(PV_Selected_Path) = True Then
            SelectedPath = UCase(PV_Selected_Path)
        Else
            SelectedPath = RemoveBackSlash(.ImageComboDir.SelectedItem.Key)
        End If
    End With
    
End Property

Function GetDriveFromPath(lsPath As String) As String

    If Len(lsPath) >= 2 Then
        GetDriveFromPath = LCase(Left(lsPath, 3))
    Else
        GetDriveFromPath = ""
    End If
    
End Function

Sub GetFolderPath(lsPath As String, lsFolder() As String)

Dim i               As Integer

    lsPath = RemoveBackSlash(lsPath, True)
    lsFolder = Split(lsPath, "\")
    For i = 1 To UBound(lsFolder)
        lsFolder(i) = lsFolder(i - 1) & "\" & lsFolder(i)
    Next

End Sub

Private Sub RemovePath(ByVal lvPath As String)

Dim lsFolderPath()  As String
Dim i               As Integer
Dim sKey            As String
Dim cboItem         As ComboItem
Dim Index, Indent   As Integer
    
    If lvPath = "" Then
        Exit Sub
    End If
    lvPath = RemoveBackSlash(lvPath)
    GetFolderPath lvPath, lsFolderPath
    With UserControl.ImageComboDir
        If .SelectedItem.Indentation >= UBound(lsFolderPath) Then
        For i = UBound(lsFolderPath) To 1 Step -1
            sKey = lsFolderPath(i) & "\"
            .ComboItems.Remove sKey
            ilsIcons16.ListImages.Remove sKey
        Next
        End If
        'cboItem.Selected = True
    End With

End Sub

Property Let SetPath(ByVal lvPath As String)

Dim lsDrive         As String
Dim lsFolder()      As String
Dim lsFolderPath()  As String
Dim i               As Integer
Dim sKey            As String
Dim cboItem         As ComboItem
Dim Index, Indent   As Integer

'   On Error GoTo SetPath_Error

    'UserControl.ImageComboDir.ComboItems(lsDrive).Selected = True
    'If UserControl.ImageComboDir.ComboItems(lsDrive) Then
    lvPath = RemoveBackSlash(lvPath)
    If PV_Selected_Path <> lvPath Then
        RemovePath PV_Selected_Path
    End If
    PV_Selected_Path = lvPath
    lsDrive = GetDriveFromPath(lvPath)
    If lsDrive <> lvPath Then
        GetFolderPath lvPath, lsFolderPath
        lsFolder = Split(lvPath, "\")
        With UserControl.ImageComboDir
            Index = .ComboItems(lsDrive).Index + 1
            Indent = .ComboItems(lsDrive).Indentation + 1
            For i = 1 To UBound(lsFolder)
                sKey = lsFolderPath(i) & "\"
                Set imgX = ilsIcons16.ListImages.Add(, sKey, GetIcon(sKey, egitSmallIcon))
                Set cboItem = .ComboItems.Add(Index, _
                                        sKey, lsFolder(i), _
                                        sKey, sKey, Indent)
                Index = Index + 1
                Indent = Indent + 1
            Next
            cboItem.Selected = True
        End With
    Else
        On Error Resume Next
        UserControl.ImageComboDir.ComboItems(lsDrive).Selected = True
        ImageComboDir_Click
        On Error GoTo 0
    End If

   On Error GoTo 0
   Exit Property

SetPath_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SetPath of Control de usuario OcxDirCboImg"
    
End Property

Private Sub DirList_Change()
FileList.Path = DirList.Path   ' Set file path.
End Sub

Private Sub DriveDir_Change()
DirList.Path = DriveDir.Drive   ' Set directory path.
End Sub

Private Sub Init_Ocx()
    
Dim lsDesktopFolder             As String
Dim lsMyPcFolder                As String

    DirList.Path = "C:\"
    'LstVwFolders.View = lvwList
    'Toolbar1.Buttons("View List").Value = 1
    GetDrives
    'Add Desktop
    lsDesktopFolder = fGetSpecialFolder(CSIDL_DESKTOP, UserControl.hWnd)
    ImageComboDir.ComboItems.Add 1, lsDesktopFolder, "Desktop", 4
    'Add the Root "My Computer" to ImageComboDir
    lsMyPcFolder = fGetSpecialFolder(CSIDL_DRIVES, UserControl.hWnd)
    ImageComboDir.ComboItems.Add 2, lsMyPcFolder, "My Computer", 1, , 1
    If lsMyPcFolder = "" Then
        AddVirtualFolder 2, CSIDL_DRIVES
    End If
    'Select "C Drive"
    ImageComboDir.ComboItems("c:\").Selected = True
    bIsDrive = True
    ImageComboDir_Click

End Sub

Private Function AddVirtualFolder(Index As Integer, CSIDL As Long)

    If UBound(PV_VirtualFolderList) < Index Then
        ReDim Preserve PV_VirtualFolderList(Index)
            PV_VirtualFolderList(Index) = CSIDL
    End If
    
End Function

Public Sub GetDrives()

Dim i                   As Integer
Dim sText, d, sKey      As String
Dim imgX                As ListImage

    For i = 0 To DriveDir.ListCount - 1
        d = DriveDir.List(i)
        sText = d
        sKey = Left$(sText, 2) & "\"
        Set imgX = ilsIcons16.ListImages.Add(, sKey, GetIcon(sKey, egitSmallIcon))
        ImageComboDir.ComboItems.Add i + 1, sKey, sText, sKey, sKey, 2
    Next
    
End Sub

Private Sub RemoveItemsIndent(Index As Integer, Indent As Integer)

    With UserControl.ImageComboDir
        Do
            If .ComboItems.Count >= Index Then
                If .ComboItems(Index).Indentation >= Indent Then
                    ilsIcons16.ListImages.Remove .ComboItems(Index).Key
                    .ComboItems.Remove Index
                Else
                    'Exit Do
                    Index = Index + 1
                End If
            Else
                Exit Do
            End If
        Loop
    End With
    
End Sub

Private Sub ImageComboDir_Click()
    
Dim txtSelFile              As String
Dim sDrive                  As String
Dim iLength                 As Integer
Dim iParentIndex            As Integer
Dim bIsDrive                As Boolean
Dim drv                     As Drive


    With UserControl.ImageComboDir
        On Error GoTo ImageComboDir_Click_Error
        txtSelFile = ""
        'On Error Resume Next
        If .SelectedItem.Key = "c:\WINDOWS\Desktop" Then
            bIsDrive = False
            GetDesktop
            RemoveItemsIndent .SelectedItem.Index + 1, 3
        ElseIf .SelectedItem.Key = "Root" Then
            'bIsDrive = False
            GetMyComputer
            RemoveItemsIndent .SelectedItem.Index + 1, 3
        ElseIf .SelectedItem.Indentation <= 2 Then
            If .SelectedItem.Key <> "" Then
                DriveDir.Drive = .SelectedItem.Key
                DirList.Path = .SelectedItem.Key
                'Get selected Item
                sDrive = .SelectedItem.Key
                iLength = Len(sDrive)
                'Store Index for AddChild
                iParentIndex = .SelectedItem.Index
                If iLength > 3 Then 'Selected Item is a Folder
                    bIsDrive = False
                    FileList.Path = .SelectedItem.Key
                Else 'Selected Item is a Drive
                    bIsDrive = True
                    DeleteChild 'Delete previous children
                    Set drv = fso.GetDrive(fso.GetDriveName(sDrive))
                    If drv.IsReady Then
                        f = drv.RootFolder
                        'Convert to lower case to avoid UpOne error
                        f = LCase(f)
                    Else
                        'MsgBox "Drive " & drv & " Not Ready", 16
                    End If
                End If
                RemoveItemsIndent .SelectedItem.Index + 1, 3
            Else
                RaiseEvent VirtualFolderSelected(PV_VirtualFolderList(.SelectedItem.Index))
                Exit Sub
            End If
        Else
            RemoveItemsIndent .SelectedItem.Index + 1, .SelectedItem.Indentation + 1
        End If
        On Error GoTo 0
        
        If bIsDrive = False Then
            PV_Selected_Path = RemoveBackSlash(.SelectedItem.Key)
        Else
            PV_Selected_Path = .SelectedItem.Key
        End If
        RaiseEvent ItemSelected(PV_Selected_Path)
        
    End With
    
   Exit Sub

ImageComboDir_Click_Error:

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImageComboDir_Click of Control de usuario OcxDirCboImg"
    On Error GoTo 0
    
End Sub

Function AddBackSlash(lsPath As String) As String

    If lsPath = "" Then
        AddBackSlash = ""
    Else
        If DoHaveBackSlash(lsPath) = False Then
            AddBackSlash = lsPath & "\"
        Else
            AddBackSlash = lsPath
        End If
    End If
    
End Function

Function DoHaveBackSlash(lsPath As String) As Boolean

    If Right$(lsPath, 1) = "\" Then
        DoHaveBackSlash = True
    Else
        DoHaveBackSlash = False
    End If

End Function

Function IsPathDrive(lsPath As String) As Boolean

    If Len(lsPath) <= 3 Then
        IsPathDrive = True
    Else
        IsPathDrive = False
    End If

End Function

Function RemoveBackSlash(lsPath As String, Optional bForce As Boolean = False) As String

    If (IsPathDrive(lsPath) = False Or bForce = True) And DoHaveBackSlash(lsPath) = True Then
        RemoveBackSlash = Left$(lsPath, Len(lsPath) - 1)
    Else
        RemoveBackSlash = lsPath
    End If
    
End Function

Private Sub LstVwFolders_AfterLabelEdit(Cancel As Integer, NewString As String)
    sKey = ImageComboDir.SelectedItem.Key
    AddNewFolder sKey, NewString
    fso.DeleteFolder sKey & "New Folder"
End Sub

Public Sub Up_One_Level()

Dim txtSelFile          As String
Dim sKey                As String
Dim sParentFolder       As String
Dim Index               As Integer

    txtSelFile = ""
    sKey = ImageComboDir.SelectedItem.Key
    Index = ImageComboDir.SelectedItem.Index
    sParentFolder = AddBackSlash(fso.GetParentFolderName(sKey))
    'code to adjust for Drive being selected
    If (Len(sParentFolder) = 3) Then
        'Parent is a Drive
        ImageComboDir.ComboItems(LCase(sParentFolder)).Selected = True
        DirList.Path = sParentFolder
        DeleteChild
        ImageComboDir_Click
    ElseIf (sParentFolder = "C:\WINDOWS\Desktop") Or Index = 1 Then
        ImageComboDir.ComboItems(1).Selected = True
        DirList.Path = sParentFolder
        DeleteChild
    ElseIf (sParentFolder = "") Then
        'No Parent, Select Desktop
        ImageComboDir.ComboItems(1).Selected = True
        DirList.Path = sParentFolder
        ImageComboDir_Click
        'GetMyComputer
    Else
        'Parent is a Folder
        ImageComboDir.ComboItems(sParentFolder).Selected = True
        DirList.Path = sParentFolder
        ImageComboDir_Click
    End If

End Sub

Public Sub AddChild(sKeyChild As String, sTextChild As String)

Dim sParent As String

sParent = fso.GetParentFolderName(sKeyChild)
iNumOfChildren = iNumOfChildren + 1
ReDim Preserve Children(iNumOfChildren)
'Add Child
ChildIndex = iParentIndex + iNumOfChildren
If Len(sParent) > 3 Then 'Indent Subdirectories
    Set ci = ImageComboDir.ComboItems.Add(ChildIndex, sKeyChild, sTextChild, 2, 3, 4)
Else
    Set ci = ImageComboDir.ComboItems.Add(ChildIndex, sKeyChild, sTextChild, 2, 3, 3)
End If
    ci.Selected = True 'Select the Item in the ImageCombo
'Store Key in Array to use for delete
Children(iNumOfChildren) = sKeyChild


End Sub

Public Sub DeleteChild()

Dim x               As Integer

    'Clear Previous children
    For x = 1 To iNumOfChildren
        sKey = Children(x)
        ImageComboDir.ComboItems.Remove (sKey)
        Children(x) = ""
    Next

    iNumOfChildren = 0
    
End Sub

Public Function CheckChild(Child As String) As Boolean

Dim sChildren           As String
Dim x                   As Integer

For x = 1 To iNumOfChildren
    sChildren = Children(x)
    If Child = sChildren Then
        CheckChild = True
        Exit Function
    Else
        CheckChild = False
    End If
Next

End Function

Private Sub CheckIcon(ByVal sFIle As String)

Dim sKey1               As String
Dim i                   As Long
Dim iHaveit             As Long
Dim imgX                As ListImage
Dim iPos                As Long
Dim itmX                As ListItem
Dim sExt                As String
Dim lSicon              As Object
Dim lLicon              As Object
Dim c                   As Integer

    ' We only want to get an icon for a given
    ' file type once, unless the file is an
    ' an executable or icon, in which case the
    ' icon is different for each instance of
    ' the extension type:

        sExt = UCase(fso.GetExtensionName(sFIle))
        If (sExt <> "EXE") And (sExt <> "ICO") And (sExt <> "DLL") And (sExt <> "OCX") And (sExt <> "HTML") And (sExt <> "LNK") And (sExt <> "") Then
            sKey1 = sExt
        Else
            sKey1 = sFIle
        End If
        sKey1 = UCase$(sKey1)
    ' Determine whether we've already got this type:
    For i = 1 To ilsIcons32.ListImages.Count
        If (ilsIcons32.ListImages(i).Key = sKey1) Then
            iHaveit = i
        End If
    Next i
    ' If we haven't already got it, then get the file
    ' icons and types and add them to the Image Lists:
    If (iHaveit = 0) Then
        
        Set imgX = ilsIcons32.ListImages.Add(, sKey1, GetIcon(sFIle, egitLargeIcon))
        imgX.Tag = GetFileTypeName(sFIle)
        iHaveit = imgX.Index
        ilsIcons16.ListImages.Add , sKey1, GetIcon(sFIle, egitSmallIcon)
        c = ilsIcons16.ListImages.Count
    End If
    
End Sub

Public Sub AddNewFolder(Path, FolderName)
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f = fso.GetFolder(Path)
  Set fc = f.SubFolders
  
  If FolderName <> "" Then
    Set nf = fc.Add(FolderName)
  Else
    Set nf = fc.Add("New Folder")
  End If
  DirList.Refresh
End Sub

Public Sub ShowFolderList()

Dim i                   As Integer
Dim iListCount          As Integer
Dim sKey, sText         As String
Dim itmFldr             As ListItem

    iListCount = LstVwFolders.ListItems.Count
    For i = 0 To DirList.ListCount - 1
        sKey = DirList.List(i)
        sText = fso.GetBaseName(sKey)
'        Set itmFldr = LstVwFolders.ListItems.Add(iListCount + i + 1, sKey, sText, 1, 2)
    Next
    
End Sub

Public Sub ShowFileList()

Dim Length, i, iListCount           As Integer
Dim ind1, ind2                      As Integer
Dim sKey, sText, sFType, sExt       As String
Dim bIsDrive                        As Boolean
Dim itmFldr                         As ListItem

    ''On Error Resume Next
    iListCount = LstVwFolders.ListItems.Count
    For i = 0 To FileList.ListCount - 1
        If bIsDrive Then
            sKey = FileList.Path & FileList.List(i)
        Else
            sKey = FileList.Path & "\" & FileList.List(i)
        End If
        sText = FileList.List(i)
        Length = Len(sText)
        'pos1 = InStr(Length - 5, sText, ".", vbBinaryCompare)
        sFType = fso.GetExtensionName(sKey) 'Mid(sText, pos1 + 1) ', 3)
        sFType = LCase(sFType)
        sExt = UCase(sFType)
        If sExt = "" Then
            sExt = "UNK"
        End If
        CheckIcon sKey
        If (sExt <> "EXE") And (sExt <> "ICO") And (sExt <> "DLL") And (sExt <> "OCX") And (sExt <> "HTML") And (sExt <> "LNK") And (sExt <> "UNK") Then
            sKey1 = sExt
        Else
            sKey1 = sKey
        End If
        sKey1 = UCase$(sKey1)
        ind1 = ilsIcons16.ListImages(sKey1).Index
        ind2 = ilsIcons32.ListImages(sKey1).Index
        Set itmFldr = LstVwFolders.ListItems.Add(iListCount + i + 1, sKey, sText, sKey1, sKey1)
    Next

End Sub

Public Sub GetlvReportData(folderspec)
'On Error Resume Next
Dim sThis As String
Dim sThat As String
    If Len(folderspec) > 3 Then
        Set f = fso.GetFolder(folderspec)
        Set fc = f.Files
    Else
        Set f = fso.GetFolder(folderspec)
        Set fc = f.SubFolders
    End If
    
'    For Each f1 In fc
'        d = f1.DateCreated
'        s = f1.Size
'        t = f1.Type
'        p = f1.Path
'        sThis = Left$(p, 1)
'        sThat = LCase$(sThis)
'        p = Replace(p, sThis, sThat, 1, 1, vbBinaryCompare)
'        Set itmX = LstVwFolders.ListItems(p)
'        itmX.SubItems(1) = s
'        itmX.SubItems(2) = t
'        itmX.SubItems(3) = d
'    Next
    
End Sub

Public Sub GetDesktop()

Dim iParentIndex            As Integer
    ''On Error Resume Next
    Screen.MousePointer = vbHourglass
'    LstVwFolders.ListItems.Clear
    DirList.Path = "C:\WINDOWS\Desktop\"
    FileList.Path = DirList.Path
    iParentIndex = ImageComboDir.SelectedItem.Index
'    ShowFolderList
'    ShowFileList
'    Label2 = LstVwFolders.ListItems.Count
    Screen.MousePointer = vbDefault

End Sub

Public Sub GetMyComputer()
    
Dim ind, i, iParentIndex        As Integer
Dim d, sText, sKey              As String

    'On Error Resume Next
    iParentIndex = ImageComboDir.SelectedItem.Index
'    LstVwFolders.ListItems.Clear
    For i = 0 To DriveDir.ListCount - 1
        d = DriveDir.List(i)
        sText = d
        sKey = Left$(sText, 2) & "\"
        ind = ilsIcons16.ListImages(sKey).Index
'        LstVwFolders.ListItems.Add i + 1, sKey, sText, ind, ind
    Next

End Sub

Private Sub UserControl_Initialize()

    ReDim PV_VirtualFolderList(0)
    Init_Ocx
    
End Sub

Private Sub UserControl_Resize()

    With UserControl
        .ImageComboDir.Width = .Width
        .Height = .ImageComboDir.Height
    End With
    
End Sub
