VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Folders"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblTempDir 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   75
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      Caption         =   "Temp Directory:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lblWinDir 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   45
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      Caption         =   "Windows Directory:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1410
   End
   Begin VB.Label lblSysDir 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      Caption         =   "System Directory:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------
' prjDirectories by Richard Gunton 12/6/00
' http://www.gunton.8k.com
'-----------------------------------------

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260

Public Function SystemDir() As String
    Dim tmp As String
    tmp = Space$(MAX_PATH)
    SystemDir = Left$(tmp, GetSystemDirectory(tmp, MAX_PATH))
End Function

Public Function Home() As String
    Dim lpBuffer As String
    lpBuffer = Space$(MAX_PATH)
    Home = Left$(lpBuffer, GetWindowsDirectory(lpBuffer, MAX_PATH))
End Function

Public Function TmpDir() As String
    Dim lpBuffer As String
    lpBuffer = Space$(MAX_PATH)
    TmpDir = Left$(lpBuffer, GetTempPath(MAX_PATH, lpBuffer))
End Function

Private Sub Form_Load()
lblSysDir.Caption = SystemDir
lblWinDir.Caption = Home
lblTempDir.Caption = TmpDir
End Sub
