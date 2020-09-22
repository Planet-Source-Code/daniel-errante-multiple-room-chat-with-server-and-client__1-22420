VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form3 
   Caption         =   "Chat - chatroomname"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   Icon            =   "chatroom.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5115
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   300
      TextRTF         =   $"chatroom.frx":0CCA
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   6720
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8070
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"chatroom.frx":0D9F
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Form2.sock.State = 7 Then
send "leav"
Form2.Show
Else
Unload Me
Load Form1
Form1.Show
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub RichTextBox2_KeyPress(KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Trim(RichTextBox2.Text) = "" Then Exit Sub
send "msg " & un$ & ": " & RichTextBox2.Text
RichTextBox2.Text = ""
End Sub
