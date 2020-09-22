VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat Room Server"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "chatroomserver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock host 
      Index           =   0
      Left            =   6480
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List3 
      Height          =   4740
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   4740
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8361
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"chatroomserver.frx":08CA
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
host(0).Close
host(0).LocalPort = 5678
host(0).Listen
msg "Host " & host(0).LocalIP & " hosting on port " & host(0).LocalPort, vbRed
chans(0).chan = "main"
chans(0).users = 0
refreshlist
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
sendall "dis "

End Sub

Private Sub host_Close(Index As Integer)
On Error Resume Next
For i = 0 To List1.ListCount - 1
    If List1.List(i) = user(Index).un Then
                remnickfromchan Index
                DoEvents
                user(Index).un = ""
                user(Index).chan = ""
                user(Index).ListIndex = 0
                List1.RemoveItem i
                msg "User " & host(Index).RemoteHostIP & " disconnected at " & Format(Time, "HH:MM:SS AM/PM"), vbBlue
    End If
Next i
        
End Sub

Private Sub host_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo hell:
Dim d As Integer
If NuM = 0 Then
NuM = 1
Load host(NuM)
host(NuM).Accept requestID
msg "User " & host(Index).RemoteHostIP & " connected at " & Format(Time, "HH:MM:SS AM/PM"), vbBlue
Exit Sub
End If
For d = 1 To NuM
    If host(d).State <> 7 Then
        host(d).Close
        host(d).Accept requestID
        msg "User " & host(d).RemoteHostIP & " connected at " & Format(Time, "HH:MM:SS AM/PM"), vbBlue
        Exit Sub
    End If
Next d
If NuM = 1000 Then Exit Sub
NuM = NuM + 1
Load host(NuM)
host(NuM).Accept requestID
msg "User " & host(Index).RemoteHostIP & " connected at " & Format(Time, "HH:MM:SS AM/PM"), vbBlue
hell:
End Sub

Private Sub host_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
strdata = ""
If host(Index).State <> 7 Then Exit Sub
host(Index).GetData strdata
cmd$ = Left(strdata, 4)
If user(Index).un = "" Then
msg host(Index).RemoteHostIP & ": " & strdata, vbBlue
Else
msg user(Index).un & ": " & strdata, vbBlue
End If
If Len(strdata) > 5 Then rest$ = Right(strdata, Len(strdata) - 4)
rest$ = Trim(rest$)
If cmd$ = "100 " Then
    send Index, "220 "
End If
If cmd$ = "user" Then
    For i = 0 To List1.ListCount - 1
    If rest$ = List1.List(i) Then
        send Index, "310 User name already in use."
        host(Index).Close
        Exit Sub
    End If
    Next i
    user(Index).un = rest$
    List1.AddItem rest$
    send Index, "300 User name accepted."
End If
If cmd$ = "chnl" Then
    Sendchans Index
End If
If cmd$ = "join" Then
    If rest$ = user(Index).chan Then
        send Index, "510 " & user(Index).un & " is already located in " & user(Index).chan & "."
        Exit Sub
    End If
    addnicktochan Index, rest$
    send Index, "500 "
End If
If cmd$ = "leav" Then
Dim t As Integer
t = Index
    remnickfromchan Index
End If
If cmd$ = "list" Then
    send Index, "lst " & chans(user(Index).ListIndex).uns
End If
If cmd$ = "msg " Then
    sendmsginchan Index, strdata
End If
If cmd$ = "cre " Then
    If addchan(rest$) = False Then
        send Index, "710 Channel already created."
    Else
        send Index, "700 Channel created."
        refreshlist
    End If
End If
If cmd$ = "snd " Then
DoEvents
    sendmsginchan Index, "600 " & rest$
End If
If cmd$ = "900 " Then
    senduserlist Index
End If
End Sub


