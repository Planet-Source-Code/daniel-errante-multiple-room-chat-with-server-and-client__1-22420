VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Join chat room"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "joinchat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sock 
      Left            =   1680
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Join"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Users: 0"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.ListIndex = -1 Then Exit Sub
If chanlist(List1.ListIndex).pw <> "" Then
    Load Form5
    Form5.Show 1, Me
    Exit Sub
End If
send "join " & chanlist(List1.ListIndex).chan
Form3.Caption = List1.List(List1.ListIndex)

End Sub

Private Sub Command2_Click()
send "chnl"

End Sub

Private Sub Command3_Click()
Load Form4
Form4.Show 1, Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End

End Sub

Private Sub List1_Click()
Label1.Caption = "Users: " & chanlist(List1.ListIndex).users

End Sub

Private Sub List1_DblClick()
Command1_Click

End Sub

Private Sub sock_Connect()
send "100 "

End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
strdata = ""
If sock.State <> 7 Then Exit Sub
sock.GetData strdata
cmd$ = Left(strdata, 4)
If Len(strdata) > 5 Then rest$ = Right(strdata, Len(strdata) - 4)
rest$ = Trim(rest$)
If cmd$ = "220-" Then
Load Form3
msg rest$, vbBlue
End If
If cmd$ = "220 " Then
send "user " & un$
End If
If cmd$ = "310 " Then
    msg rest$, vbRed
    Form3.Show
End If
If cmd$ = "300 " Then
send "chnl"
    Me.Show
    Unload Form1
End If
If cmd$ = "400 " Then
temp = List1.ListIndex
    addchans rest$
On Error GoTo hell:
List1.Selected(temp) = True
hell:
End If
If cmd$ = "500 " Then
    Unload Form4
    Form3.Show
    Me.Hide
    send "snd " & un$ & " has joined the channel."
End If

If cmd$ = "lst " Then
Form3.List1.Clear
    For i = 1 To Len(rest$)
        m$ = Mid(rest$, i, 1)
            If m$ = ";" Then
                Form3.List1.AddItem t$
                t$ = ""
            Else
            t$ = t$ & m$
            End If
    Next i
End If
If cmd$ = "msg " Then
pos% = InStr(1, rest$, ":")
    If Left(rest$, pos% - 1) = un$ Then
        msg rest$, vbBlue
    Else
        msg rest$
    End If
End If
If cmd$ = "700 " Then
If Form4.Text2.Text = "" Then
    Form3.Caption = Form4.Text1.Text
    Else
    Form3.Caption = Form4.Text1.Text & "*"
End If
    send "join " & Form4.Text1.Text
End If
If cmd$ = "710 " Then
    Form4.Label4.Caption = rest$
    Form4.Text1.Enabled = True
    Form4.Text2.Enabled = True
    Form4.Text3.Enabled = True
End If
If cmd$ = "600 " Then
    msg rest$, vbRed
    send "900 "
End If
If cmd$ = "dis " Then
MsgBox "Disconnected from server!", vbExclamation, "Disconnected!"
sock.Close
Load Form1
Form1.Show
Unload Form5
Unload Form4
Unload Form3
End If
If cmd$ = "800 " Then
    send "chnl"
End If
End Sub

