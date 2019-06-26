VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInstanceControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mForm       As Form
Private mLoaded     As Boolean
Private mLoad       As Boolean

Sub SetLoaded()
    mLoaded = True
    mLoad = False
End Sub

Sub ClearLoaded()
    mLoaded = False
End Sub

Function GetLoaded() As Boolean
    GetLoaded = mLoaded
End Function

Sub SetLoad()
    mLoad = True
End Sub

Sub ClearLoad()
    mLoad = False
End Sub

Function GetLoad() As Boolean
    GetLoad = mLoad
End Function

