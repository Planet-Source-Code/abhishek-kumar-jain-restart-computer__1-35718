VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Restart Your Computer Now..."
      Default         =   -1  'True
      Height          =   435
      Left            =   345
      TabIndex        =   0
      Top             =   720
      Width           =   4065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Private Const SystemChangeRestart = 4

Public Sub SettingsChanged(FormName As Form)
    SHRestartSystemMB FormName.hWnd, vbNullString, SystemChangeRestart
End Sub

Private Sub Command1_Click()
    SettingsChanged Me
End Sub
