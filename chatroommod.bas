Attribute VB_Name = "Module1"
Public un$
Public NuM2 As Integer

Type chanlist1
chan As String
pw As String
users As Integer
End Type

Public chanlist(1000) As chanlist1

Function msg(message As String, Optional clr As String = vbBlack, Optional bld As Boolean = False, Optional sze As Long = 10)
Form3.RichTextBox1.SelStart = Len(Form3.RichTextBox1.Text)
Form3.RichTextBox1.SelColor = clr
Form3.RichTextBox1.SelFontSize = sze
Form3.RichTextBox1.SelBold = bld
Form3.RichTextBox1.SelText = message & vbCrLf
Form3.RichTextBox1.SelStart = Len(Form3.RichTextBox1.Text)

End Function

Function send(dat As String)
If Form2.sock.State = 7 Then Form2.sock.SendData dat
If Form2.sock.State = 8 Then Form2.sock.Close
DoEvents
End Function

Function addchans(thetext As String)
Dim d As Integer
Form2.List1.Clear
For i = 1 To Len(thetext)
    m$ = Mid(thetext, i, 1)
        If m$ = "," Then
            chanlist(d).users = t$
            t$ = ""
        ElseIf m$ = ";" Then
            chanlist(d).chan = t$
            t$ = ""
        ElseIf m$ = ">" Then
            chanlist(d).pw = t$
            If chanlist(d).pw <> "" Then
                Form2.List1.AddItem chanlist(d).chan & "*"
            Else
                Form2.List1.AddItem chanlist(d).chan
            End If
            d = d + 1
            NuM2 = d
            t$ = ""
        Else
        t$ = t$ & m$
        End If
Next i
End Function
