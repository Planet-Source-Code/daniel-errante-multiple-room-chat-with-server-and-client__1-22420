VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Channel"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "createhim.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   1
      Top             =   240
      Width           =   2500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Chat room name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Text <> Text3.Text Then GoTo hell:
If Text1.Text = "" Then GoTo hell:
If InStr(1, Text1.Text, ">") <> 0 Then GoTo hell:
If InStr(1, Text1.Text, ",") <> 0 Then GoTo hell:
If InStr(1, Text1.Text, ";") <> 0 Then GoTo hell:
If InStr(1, Text1.Text, "*") <> 0 Then GoTo hell:
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Label4.Caption = "Contacting server..."
send "cre " & Text1.Text & "," & Text2.Text
Exit Sub
hell:
MsgBox "If you specified a password, make sure they are the same, and that the room name does not contain the characters '>' ',' or ';'", vbExclamation
End Sub

Private Sub Command2_Click()
Unload Me

End Sub
