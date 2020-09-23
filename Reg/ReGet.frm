VERSION 5.00
Begin VB.Form ReGet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Made By Nir Schwartz"
   ClientHeight    =   4050
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cache"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Programs"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start Up"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Menu"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Desktop"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Menu mnuString 
      Caption         =   "String"
      Visible         =   0   'False
      Begin VB.Menu mnuNewString 
         Caption         =   "New String"
      End
      Begin VB.Menu mnuDelString 
         Caption         =   "Delete String"
      End
   End
   Begin VB.Menu mnuKey 
      Caption         =   "Key"
      Visible         =   0   'False
      Begin VB.Menu mnuNewKey 
         Caption         =   "New Key"
      End
      Begin VB.Menu mnuDelKey 
         Caption         =   "Delete Key"
      End
   End
End
Attribute VB_Name = "ReGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
MsgBox (GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop"))
End Sub

Private Sub Command2_Click()
MsgBox (GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Start Menu"))

End Sub

Private Sub Command3_Click()
MsgBox (GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Startup"))

End Sub

Private Sub Command4_Click()
MsgBox (GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs"))
End Sub

Private Sub Command5_Click()
MsgBox (GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cache"))
End Sub

Private Sub Command6_Click()
End
End Sub
