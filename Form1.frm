VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Xat Client Plus Alpha 0.9              -InmortalHS-"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   17
      Text            =   "634545098"
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Text            =   "356016616"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Limpiar Ids"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Kickeo Múltiple >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Text            =   "244856312"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Text            =   "610320264"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Text            =   "415075241"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Text            =   "259847792"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Text            =   "262865723"
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Text            =   "160404399"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Text            =   "414854557"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Text            =   "69184570"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cargar ids >>>"
      Height          =   615
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Baneo Multiple >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   450
      Left            =   2640
      Top             =   4800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Flooding Test"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reconectar Socket"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   720
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Ids"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Host As String = "m.xat.com"
Private Sub Command1_Click()
Winsock1.Close
Winsock1.Connect Host, 10049

'Winsock1.RemoteHost = "m.xat.com"
'Winsock1.RemotePort = "10049"
End Sub

Private Sub Command2_Click()
Dim x As Integer
Dim Peticion
For x = 0 To 10

Peticion = "GET /Post?m=Probando" & x & " HTTP/1.1" & vbCrLf & _
            "Host: " & Host & vbCrLf & _
            "Connection: close" & vbCrLf & vbCrLf
Winsock1.SendData Peticion
Pause (0.8) ' Intervalo mínimo que permite xat, no cambiar

Next x


End Sub

Private Sub Command3_Click()

Dim x As Integer
Dim Peticion
For x = 0 To List1.ListCount - 1

Peticion = "GET /Post?p=test" & List1.List(x) & "&u=" & List1.List(x) & "&t=/g1 HTTP/1.1" & vbCrLf & _
            "Host: " & Host & vbCrLf & _
            "Connection: close" & vbCrLf & vbCrLf
            
Winsock1.SendData Peticion
Pause (0.9) ' Intervalo mínimo que permitido, no cambiar

Next x

End Sub

Private Sub Command4_Click()

List1.AddItem Text1(0)
List1.AddItem Text1(1)
List1.AddItem Text1(2)
List1.AddItem Text1(3)
List1.AddItem Text1(4)
List1.AddItem Text1(5)
List1.AddItem Text1(6)
List1.AddItem Text1(7)
List1.AddItem Text1(8)
List1.AddItem Text1(9)
Command3.Enabled = True
Command5.Enabled = True

End Sub

Private Sub Command5_Click()
Dim x As Integer
Dim Peticion
For x = 0 To List1.ListCount - 1

Peticion = "GET /Post?p=test" & List1.List(x) & "&u=" & List1.List(x) & "&t=/k HTTP/1.1" & vbCrLf & _
            "Host: " & Host & vbCrLf & _
            "Connection: close" & vbCrLf & vbCrLf
            
Winsock1.SendData Peticion
Pause (0.8) ' Intervalo mínimo que permitido, no cambiar

Next x
End Sub

Private Sub Command6_Click()
List1.Clear
End Sub



Private Sub Timer1_Timer()
If Winsock1.State = sckConnected Then
    Shape2.Visible = True
Shape1.Visible = False
Else
    Shape2.Visible = False
Shape1.Visible = True
End If
End Sub


Private Sub Winsock1_SendComplete()
Winsock1.Close
Winsock1.Connect Host, 10049
End Sub
