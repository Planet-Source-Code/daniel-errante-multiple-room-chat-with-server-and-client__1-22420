Attribute VB_Name = "chatserv"
Public NuM As Integer
Public NuM2 As Integer

Type user1
chan As String
ListIndex As Integer
un As String
End Type

Type chans1
chan As String
users As Integer
uns As String
pw As String
End Type

Public chans(1000) As chans1
Public user(1000) As user1

Function msg(message As String, Optional clr As String = vbBlack, Optional bld As Boolean = False, Optional sze As Long = 10)
Form1.RichTextBox1.SelStart = Len(Form1.RichTextBox1.Text)
Form1.RichTextBox1.SelColor = clr
Form1.RichTextBox1.SelFontSize = sze
Form1.RichTextBox1.SelBold = bld
Form1.RichTextBox1.SelText = message & vbCrLf
End Function

Function send(prt As Integer, dat As String, Optional showmsg As Boolean = True)
If Form1.host(prt).State = 7 Then Form1.host(prt).SendData dat
If Form1.host(prt).State = 8 Then Form1.host(prt).Close
stri$ = dat
If showmsg = True Then msg stri$, vbRed
DoEvents
End Function

Function Sendchans(prt2 As Integer)
For i = 0 To NuM2
If chans(i).chan <> "" Then m$ = m$ & chans(i).users & "," & chans(i).chan & ";" & chans(i).pw & ">"
Next i
send prt2, "400 " & m$

End Function

Function addnicktochan(ind As Integer, thechan As String)
If user(ind).chan = "" Then GoTo nxt2:
For i = 1 To Len(chans(user(ind).ListIndex).uns)
    m$ = Mid(chans(user(ind).ListIndex).uns, i, 1)
    If m$ = ";" Then
        If user(ind).un = t$ Then
            chans(user(ind).ListIndex).uns = Left(chans(user(ind).ListIndex).uns, i - Len(user(ind).un)) & Right(chans(user(ind).ListIndex).uns, Len(chans(user(ind).ListIndex).uns) - Len(user(ind).un) - i)
            GoTo nxt2:
        End If
    Else
    t$ = t$ & m$
    End If
Next i
nxt2:
For i = 0 To Form1.List2.ListCount - 1
    If thechan = Form1.List2.List(i) Then
        user(ind).ListIndex = i
        GoTo nxt:
    End If
Next i
nxt:
chans(user(ind).ListIndex).uns = chans(user(ind).ListIndex).uns & user(ind).un & ";"
chans(user(ind).ListIndex).users = chans(user(ind).ListIndex).users + 1
refreshlist
user(ind).chan = thechan
End Function

Function refreshlist()
Form1.List2.Clear
Form1.List3.Clear
For i = 0 To NuM2
    If chans(i).chan <> "" Then
    Form1.List2.AddItem chans(i).chan
    Form1.List3.AddItem chans(i).users
    End If
Next i
End Function

Function remnickfromchan(ind As Integer)
If user(ind).chan = "" Then GoTo nxt2:
For i = 1 To Len(chans(user(ind).ListIndex).uns)
    m$ = Mid(chans(user(ind).ListIndex).uns, i, 1)
    If m$ = ";" Then
    tt$ = t$
    t$ = ""
        If user(ind).un = tt$ Then
        Dim uns As String
        uns = chans(user(ind).ListIndex).uns
            uns = Left(uns, i - Len(user(ind).un) - 1) & Right(uns, Len(uns) - i)
            chans(user(ind).ListIndex).uns = uns
            chans(user(ind).ListIndex).users = chans(user(ind).ListIndex).users - 1
            refreshlist
            If chans(user(ind).ListIndex).users <> 0 Then senduserlist ind
            user(ind).chan = ""
            If chans(user(ind).ListIndex).users = 0 Then
                If chans(user(ind).ListIndex).chan <> "main" Then
                    remchan (chans(user(ind).ListIndex).chan)
                End If
            End If
            GoTo nxt2:
        End If
    Else
    t$ = t$ & m$
    End If
Next i
nxt2:
send ind, "800 "
refreshlist
End Function

Function sendmsginchan(theind As Integer, thedat As String)
Dim i As Integer
For i = 1 To NuM
    If Form1.host(i).State = 7 Then
    If user(i).chan = user(theind).chan Then send i, thedat, False
    End If
Next i
End Function


Function senduserlist(userind As Integer)
Dim i As Integer
For i = 1 To NuM
    If Form1.host(i).State = 7 Then
    If user(i).chan = user(userind).chan Then send i, "lst " & chans(user(userind).ListIndex).uns, False
    End If
Next i
End Function

Function sendall(thedat As String)
Dim i As Integer
For i = 1 To NuM
    If Form1.host(i).State = 7 Then send i, thedat
Next i
End Function

Function addchan(thename As String) As Boolean
For i = 1 To Len(thename)
m$ = Mid(thename, i, 1)
    If m$ = "," Then
        room$ = t$
        pw$ = Right(thename, Len(thename) - Len(room$) - 1)
        GoTo nxt:
    Else
    t$ = t$ & m$
    End If
Next i
nxt:
For i = 0 To NuM2
    If chans(i).chan = room$ Then
        addchan = False
        Exit Function
    End If
Next i
For i = 0 To NuM2
    If chans(i).chan = "" Then
        chans(i).chan = room$
        chans(NuM2).users = 0
        If pw$ <> "" Then chans(i).pw = pw$
        addchan = True
        Exit Function
    End If
Next i
NuM2 = NuM2 + 1
chans(NuM2).chan = room$
chans(NuM2).users = 0
If pw$ <> "" Then chans(NuM2).pw = pw$
addchan = True
End Function

Function remchan(thename As String)
For i = 0 To NuM2
If thename = chans(i).chan Then
    chans(i).chan = ""
    chans(i).uns = ""
    chans(i).users = 0
    chans(i).pw = ""
End If
Next i

End Function
