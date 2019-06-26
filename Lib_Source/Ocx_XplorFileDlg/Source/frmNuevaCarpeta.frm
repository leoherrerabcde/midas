VERSION 5.00
Begin VB.Form frmNuevaCarpeta 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdCreateFolder 
      Caption         =   "Crear Carpeta"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtNewFolder 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "txtNewFolder"
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblNewName 
      Caption         =   "Especifique el Nombre de la Nueva Carpeta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmNuevaCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmNuevaCarpeta
' Author    : Leo Herrera
' Date      : 14/10/2013
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private PV_Last_Path                As String


Sub Set_Last_Path(lvPath As String)

End Sub
