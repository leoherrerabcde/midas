VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Declare Function FindFirstFile Lib "kernel32" Alias _
        "FindFirstFileA" (ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindNextFile Lib "kernel32" Alias _
        "FindNextFileA" (ByVal hFindFile As Long, _
        lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindClose Lib "kernel32" (ByVal _
        hFindFile As Long) As Long
        
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias _
        "SHGetFileInfoA" (ByVal pszPath As String, ByVal _
        dwFileAttributes As Long, psfi As ShellFileInfoType, ByVal _
        cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" _
        (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) _
        As Long
        
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" _
        (pDicDesc As IconType, riid As CLSIdType, ByVal fown As Long, _
        lpUnk As Object) As Long
        

Const SHGFI_TYPENAME = &H400&
Const MAX_PATH = 259
Const Large = &H100
Const Small = &H101

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Type IconType
  cbSize As Long
  picType As PictureTypeConstants
  hIcon As Long
End Type

Private Type CLSIdType
  id(16) As Byte
End Type

Private Type ShellFileInfoType
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type

Private Type FileType
  Name As String
  Ext As String
  IcoIndex As Integer
End Type

Private Type DIRLISTTYPE
  Ext As String
  File As String
  Type As String
  Attributes As Long
  FileLen As Long
  LastWrite As FILETIME
End Type

Dim DirList() As DIRLISTTYPE

Private mvarImageList1 As ImageList
Private mvarImageList2 As ImageList

Private mvarPath As String


Private WithEvents mvarListView As ListView
Attribute mvarListView.VB_VarHelpID = -1
Public Event changePath(Ruta As String)

Public Property Let Path(ByVal vData As String)
    mvarPath = vData
End Property


Public Property Get Path() As String
    Path = mvarPath
End Property

Public Function SetControls( _
    ByRef ListView As ListView, _
    ImageList1 As ImageList, _
    ImageList2 As ImageList)
    
    Set mvarListView = ListView
    
    Set mvarImageList1 = ImageList1
    Set mvarImageList2 = ImageList2
    
    mvarImageList1.BackColor = mvarListView.BackColor
    mvarImageList2.BackColor = mvarListView.BackColor
    
    mvarImageList1.MaskColor = mvarListView.BackColor
    mvarImageList2.MaskColor = mvarListView.BackColor
    
    
    Iniciar_ListView
End Function


Public Sub Listar(ByVal Path As String)
  Dim tItem As ListItem, X%, Y&, Ext$
  
    mvarPath = Path
    
    mvarListView.ListItems.Clear
    
    If Not GetFiles(Path, "*.*") Then Exit Sub
    
    DoEvents
      
    Set mvarListView.Icons = mvarImageList2
    Set mvarListView.SmallIcons = mvarImageList1
       
    For X = 0 To UBound(DirList) - 1
      Ext = DirList(X).Ext
      If Ext <> "" Then
        Set tItem = mvarListView.ListItems.Add(, , DirList(X).File, Ext, Ext)
                                          
        If Ext <> "carpeta" Then
          Y = DirList(X).FileLen \ 1024 + 1
          tItem.SubItems(1) = Y & " KB"
        End If
      End If

      tItem.SubItems(2) = DirList(X).Type
      tItem.SubItems(3) = CalcFTime(DirList(X).LastWrite)
      tItem.SubItems(4) = GetAttributes(DirList(X).Attributes)
      If X Mod 1000 = 0 Then mvarListView.Refresh
    Next X
End Sub

Private Function GetFiles(Root As String, ex As String) As Boolean
  Dim File$, hFile&, FD As WIN32_FIND_DATA
  Dim Ext$, Lcnt%, X&, Y&, Extr%, carpeta As Boolean
  Dim Descrp As String, DescrpCol As String
  
  Const Extra = "|zico=Symbol|zexe=Aplicación|" & _
                "zlnk=Verknüpfung|zcur=Cursor|"

    If right$(Root, 1) <> "\" And right$(Root, 1) <> "/" Then
      Root = Root & "\"
    End If

    DescrpCol = "|carpeta=Carpeta de Archivos|"
    ReDim DirList(0)
    Set mvarListView.Icons = Nothing
    Set mvarListView.SmallIcons = Nothing
    mvarImageList1.ListImages.Clear
    mvarImageList2.ListImages.Clear
  
    hFile = FindFirstFile(Root & ex, FD)
    If hFile = 0 Then Exit Function
    Do
       File = left(FD.cFileName, InStr(FD.cFileName, Chr(0)) - 1)
       If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
         = FILE_ATTRIBUTE_DIRECTORY Then
         If (File <> ".") And (File <> "..") Then
           With DirList(UBound(DirList))
             .Attributes = FD.dwFileAttributes
             .FileLen = FD.nFileSizeLow
             .LastWrite = FD.ftLastWriteTime
             .Ext = "carpeta"
             .File = File
             .Type = "Carpeta de Archivos"
           End With
           ReDim Preserve DirList(0 To UBound(DirList) + 1)
           If Not carpeta Then carpeta = True
         End If
       Else
         Ext = "z" & LCase(GetExtension(File))
         
         X = InStr(1, DescrpCol, "|" & Ext & "=")
         Y = InStr(1, Extra, "|" & Ext & "=")
         
         If X = 0 Or Y <> 0 Then
           If Y <> 0 Then
             Y = Y + Len(Ext) + 2
             X = InStr(Y, Extra, "|")
             Descrp = Mid$(Extra, Y, X - Y)
             
             Ext = Ext + CStr(Extr)
             Extr = Extr + 1
           Else
             Descrp = GetFileType(Root & File)
             DescrpCol = DescrpCol & Ext & "=" & Descrp & "|"
           End If
           mvarImageList1.ListImages.Add , Ext, LoadIcon(Small, Root & File)
           mvarImageList2.ListImages.Add , Ext, LoadIcon(Large, Root & File)
           Lcnt = Lcnt + 1
         Else
           X = X + Len(Ext) + 2
           Y = InStr(X, DescrpCol, "|")
           Descrp = Mid$(DescrpCol, X, Y - X)
         End If
         
         With DirList(UBound(DirList))
           .Attributes = FD.dwFileAttributes
           .FileLen = FD.nFileSizeLow
           .LastWrite = FD.ftLastWriteTime
           .Ext = Ext
           .File = File
           .Type = Descrp
         End With
         ReDim Preserve DirList(0 To UBound(DirList) + 1)
       End If
    Loop While FindNextFile(hFile, FD)
    Call FindClose(hFile)
    
    If carpeta Then
      mvarImageList1.ListImages.Add , "carpeta", LoadIcon(Small, App.Path)
      mvarImageList2.ListImages.Add , "carpeta", LoadIcon(Large, App.Path)
    End If
    
    If UBound(DirList) > 0 Then
      Call SortName
      GetFiles = True
    End If
End Function

Private Function GetExtension(ByVal FileName$) As String
  Dim aa$, BB$, X&
    For X = Len(FileName) To 1 Step -1
      If Mid$(FileName, X, 1) = "." Then Exit For
    Next X
    GetExtension = Mid$(FileName, X + 1)
End Function

Private Function LoadIcon(Size As Long, File As String) As IPictureDisp
  Dim Result&, Unkown As IUnknown, Icon As IconType, CLSID As CLSIdType
  Dim ShellInfo As ShellFileInfoType

    Call SHGetFileInfo(File, 0, ShellInfo, Len(ShellInfo), Size)
    
    Icon.cbSize = Len(Icon)
    Icon.picType = vbPicTypeIcon
    Icon.hIcon = ShellInfo.hIcon
    CLSID.id(8) = &HC0
    CLSID.id(15) = &H46
    Result = OleCreatePictureIndirect(Icon, CLSID, 1, Unkown)
    Set LoadIcon = Unkown
End Function

Private Function GetFileType(File$) As String
  Dim Result&
  Dim ShellInfo As ShellFileInfoType
    
    Call SHGetFileInfo(File, 0, ShellInfo, Len(ShellInfo), _
                       SHGFI_TYPENAME)
                       
    Result = InStr(1, ShellInfo.szTypeName, Chr$(0)) - 1
    GetFileType = left(ShellInfo.szTypeName, Result)
End Function

Private Function GetAttributes(Att&) As String
  Dim aa$
  
    If Att And FILE_ATTRIBUTE_TEMPORARY Then aa = "T"
    If Att And FILE_ATTRIBUTE_READONLY Then aa = aa & "R"
    If Att And FILE_ATTRIBUTE_HIDDEN Then aa = aa & "H"
    If Att And FILE_ATTRIBUTE_SYSTEM Then aa = aa & "S"
    If Att And FILE_ATTRIBUTE_NORMAL Then aa = aa & "N"
    If Att And FILE_ATTRIBUTE_COMPRESSED Then aa = aa & "C"
    If Att And FILE_ATTRIBUTE_ARCHIVE Then aa = aa & "A"
    GetAttributes = aa
End Function


Private Function CalcFTime(FTime As FILETIME) As String
  Dim Datum$, Zeit$, aa$, hh$, mm$, ss$, DT As Date
  Dim Da$, Mo$, Ye$
  Dim STime As SYSTEMTIME
  
    Call FileTimeToSystemTime(FTime, STime)
    With STime
      Da = .wDay
      If Len(Da) < 2 Then Da = "0" & Da
      
      Mo = .wMonth
      If Len(Mo) < 2 Then Mo = "0" & Mo
      
      Ye = CStr(.wYear)
      
      mm = Trim$(CStr(.wMinute))
      If Len(mm) < 2 Then mm = "0" & mm
      
      ss = Trim$(CStr(.wSecond))
      If Len(ss) < 2 Then ss = "0" & ss
      
      hh = Trim$(CStr(.wHour))
      If Len(.wHour) < 2 Then hh = "0" & hh

      CalcFTime = Da & "." & Mo & "." & Ye & " " _
                  & hh & ":" & mm & ":" & ss
    End With
End Function

Private Function SortName()
  Dim X%, Mem As DIRLISTTYPE, Max%
  
   For X = 0 To UBound(DirList) - 1
     If DirList(X).Type = "Carpeta de Archivos" Then
       Mem = DirList(Max)
       DirList(Max) = DirList(X)
       DirList(X) = Mem
       Max = Max + 1
     End If
   Next X
   
   If Max <> 0 Then Call QuickSort(0, Max - 1)
   If Max < UBound(DirList) Then
     Call QuickSort(Max, UBound(DirList) - 1)
   End If
End Function

Private Sub QuickSort(ByVal LB As Long, ByVal UB As Long)
  Dim P1&, P2&, Ref$, TEMP As DIRLISTTYPE

    P1 = LB
    P2 = UB
    Ref = DirList((P1 + P2) / 2).File
    
    Do
      Do While (DirList(P1).File < Ref)
        P1 = P1 + 1
      Loop
 
      Do While (DirList(P2).File > Ref)
        P2 = P2 - 1
      Loop

      If P1 <= P2 Then
        TEMP = DirList(P1)
        DirList(P1) = DirList(P2)
        DirList(P2) = TEMP
            
        P1 = P1 + 1
        P2 = P2 - 1
      End If
    Loop Until (P1 > P2)

    If LB < P2 Then Call QuickSort(LB, P2)
    If P1 < UB Then Call QuickSort(P1, UB)
End Sub


Function subirNivel()
    
    Dim spath As String
    spath = mvarPath
    
    If Len(spath) > 3 Then
    
        If right(spath, 1) = "\" Then
            spath = left(spath, Len(spath) - 1)
        End If
        spath = left(spath, Len(spath) - Len(right(spath, Len(spath) - InStrRev(spath, "\"))))
        
        Me.Listar spath
        RaiseEvent changePath(mvarPath)
    End If
    
End Function

Private Sub mvarListView_DblClick()
  Dim X&, aa$
  
    X = mvarListView.SelectedItem.Index
    If X <> 0 Then
      If DirList(X - 1).Type = "Carpeta de Archivos" Then
        aa = mvarPath
        If right$(aa, 1) <> "\" And right$(aa, 1) <> "/" Then
          aa = aa & "\"
        End If
        Listar aa & DirList(X - 1).File
        RaiseEvent changePath(mvarPath)
      End If
    End If
End Sub

Private Sub Iniciar_ListView()
  Dim TPX  As Long
    TPX = Screen.TwipsPerPixelX
    With mvarListView
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Nombre", 140 * TPX, 0
        .ColumnHeaders.Add , , "Tamaño", 70 * TPX, 1
        .ColumnHeaders.Add , , "Tipo", 160 * TPX, 0
        .ColumnHeaders.Add , , "Fecha", 120 * TPX, 0
        .ColumnHeaders.Add , , "Atributo", 60 * TPX, 1
    End With
End Sub

Private Sub Class_Terminate()
    Set mvarImageList1 = Nothing
    Set mvarImageList2 = Nothing
    Set mvarListView = Nothing
End Sub




