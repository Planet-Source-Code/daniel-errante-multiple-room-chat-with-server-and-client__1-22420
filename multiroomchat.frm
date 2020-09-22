VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "multiroomchat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1440
      Top             =   960
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "User name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
rep$ = InputBox$("Enter the IP address of the server you want to add:", "New Server")
If rep$ = "" Then
Exit Sub
Else
For i = 0 To Combo1.ListCount - 1
If Combo1.List(i) = rep$ Then Exit Sub
Next i
Combo1.AddItem rep$
End If
End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Command3_Click()
If Combo1.Text = "retrieve from web" Then
Timer1.Enabled = True
        Combo1.Enabled = False
        Text1.Enabled = False
        Me.Caption = "Retrieving server address..."
        If Inet1.StillExecuting = False Then
        t$ = Inet1.OpenURL("http://www.hayesproductionsinc.com/server.txt")
        Else
        MsgBox "The server's address has not been found yet.  Please exit this program and restart it to try again.", vbExclamation
        End If
        For i = 0 To Combo1.ListCount - 1
            If Combo1.List(i) = "retrieve from web" Then
                Combo1.RemoveItem i
                Combo1.AddItem t$, 0
                Combo1.ListIndex = 0
                GoTo nxt:
            End If
        Next i
nxt:
        Combo1.Enabled = True
        Text1.Enabled = True
        Me.Caption = "Login"
End If
THEFILE = App.Path & "\dchat.ini"
WriteIni "settings", "user", Text1.Text
WriteIni "servers", "#", Combo1.ListCount
For i = 0 To Combo1.ListCount - 1
    For d = 0 To Combo1.ListCount - 1
    If ReadIni2("servers", "server" & d + 1) = Combo1.List(i) Then
    GoTo nxt2:
    End If
    Next d
    If Combo1.List(i) = "" Then GoTo nxt2:
    WriteIni "servers", "server" & i + 1, Combo1.List(i)
nxt2:
Next i
WriteIni "servers", "default", Combo1.ListIndex
un$ = Text1.Text
Form2.sock.Close
Form2.sock.RemoteHost = Combo1.List(Combo1.ListIndex)
Form2.sock.RemotePort = 5678
Form2.sock.Connect
End Sub

Private Sub Command4_Click()
On Error GoTo hell:
If Combo1.List(Combo1.ListIndex) = "retrieve from web" Then Exit Sub
Combo1.RemoveItem Combo1.ListIndex
hell:

End Sub

Private Sub Form_Load()

On Error GoTo hell:
THEFILE = App.Path & "\dchat.ini"
Combo1.AddItem "retrieve from web"
Text1.Text = ReadIni2("settings", "user")
For i = 1 To ReadIni2("servers", "#")
Combo1.AddItem ReadIni2("servers", "server" & i)
Next i
Combo1.ListIndex = ReadIni2("servers", "default")
Load Form2
hell:
End Sub

Private Sub Timer1_Timer()
If Inet1.StillExecuting = True Then
Inet1.Cancel
Combo1.Enabled = True
Text1.Enabled = True
Me.Caption = "Login"
End If
Timer1.Enabled = False
End Sub
