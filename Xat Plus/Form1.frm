VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Xat Client Plus Alpha 0.9          -InmortalHS-"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "153256001"
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "tupass"
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "inmortalhs"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acceder"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   6360
      Width           =   255
      ExtentX         =   450
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "ID"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ie As Object





Const Host As String = "m.xat.com"
Private Sub Command1_Click()
Form1.Winsock1.Close
Form1.Winsock1.Connect Host, 10049
WebBrowser1.Document.getElementByid("YourEmail").Value = Text1.Text
WebBrowser1.Document.getElementByid("password").Value = Text2.Text
WebBrowser1.Document.getElementByid("Group").Value = Text3.Text
WebBrowser1.Document.getElementByid("SignIn").Click
Form1.Show
Form2.Hide

End Sub




'Private Sub Command2_Click()
'Dim x As Integer

'x = 1

'Do While x <= 10
'Set ie = CreateObject("InternetExplorer.Application")
'ie.Visible = False
'ie.Navigate "http://m.xat.com:10049/Post?m=" & x & ""




'x = x + 1

'Loop
'End Sub



Private Sub Form_Load()

WebBrowser1.Navigate "http://m.xat.com/"
WebBrowser1.Silent = True

End Sub


Private Sub WebBrowser1_DownloadComplete()
Command1.Enabled = True
End Sub

