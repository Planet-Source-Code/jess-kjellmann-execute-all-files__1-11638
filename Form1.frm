VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Run File"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   3135
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Thank you for downloading this code.
' Many users uses SHELL, but that command can't execute all files
' but this can
' You can use the code in your own app, but I would like if you would
' put my name someplace in your app... Or at least vote for me :o)
' Greetings from McGoat
' Sorry about the bad english, I'm from Denmark... :o)

Option Explicit ' Does so all variables must be defined
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
ShellExecute hwnd, "open", File1.filename, "", "", vbNormalFocus
' Shellexecute hwnd , "command" , file , "parameters" , "directory" , showcmd
' command:
' open, print, explore
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path ' Make file1 dir change to Dir1 dir
ChDir Dir1.Path ' Make system DirPath to Dir1 Path
End Sub

Private Sub Drive1_Change()
Dim Answer As ButtonConstants ' Make Answer a variable
On Error GoTo erro ' If an error comes goto erro:
Dir1.Path = Drive1.Drive ' Make Dir1 dir change to Drive1 Drive
ChDrive Drive1.Drive ' Make system DrivePath to Drive1 Drive
Exit Sub ' If no errors then exit sub
erro: ' An error has come
Answer = MsgBox("Can't Access drive, try again?", vbOKCancel, "Error") ' ask user for accessing drive again
If Answer = vbOK Then Drive1_Change ' Is user press Ok, then try again
End Sub
