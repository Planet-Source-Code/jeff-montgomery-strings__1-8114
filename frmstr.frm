VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   1815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRevers 
      Caption         =   "Reverse String"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Cmducase 
      Caption         =   "Get Uppercase"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdlcase 
      Caption         =   "Get Lowecase"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdlen 
      Caption         =   "Get Length"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Cmdinstr 
      Caption         =   "Get Instr"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Cmdmid 
      Caption         =   "Get Middle"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Cmdgetrgt 
      Caption         =   "Get Right"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtlft 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Project1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Cmdgetlft 
      Caption         =   "Get left"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Im not the best programmer and i
'dont claim to be.  I did the best
'with what i know.  I hope
'this example helps you understand
'a little more about strings.
'If you have any questions or
'problems with this example
'email me at Pacex@hotmail.com



Private Sub Cmdgetlft_Click()
Dim txt$, lftstr$
txt$ = txtlft.Text
lftstr$ = Left(txt$, 3)
MsgBox lftstr$, vbOKOnly, "First 3 letters"
'this will get the 3 first letters of the
'string.  The string i am using is
'txt$ which is txtlft.text.
'Left(txt$,3) 3 is the characters to the left i
'want to get. So since my string is "Project1"
'it will return "Pro"
End Sub

Private Sub Cmdgetrgt_Click()
Dim txt$, Rightstr$
txt$ = txtlft.Text
Rightstr$ = Right(txt$, 3)
MsgBox Rightstr$, vbOKOnly, "Last 3 letters"
'this will get the 3 last letters of the string
'Right function is the exact opposite of left
'so since my string is "Project1" it will
'return "ct1"
End Sub

Private Sub Cmdinstr_Click()
Dim txt$, instring$
txt$ = txtlft.Text
instring$ = InStr(txt$, "e")
MsgBox instring$, vbOKOnly, "E"
'This will return the position of "e"
'Since e is the 5th chatacter it
'will return "5" which is the position of "e"
End Sub

Private Sub cmdlen_Click()
Dim txt$, length$
txt$ = txtlft.Text
length$ = Len(txt$)
MsgBox length$, vbOKOnly, "Length"
'This will return the lenth of txt$
End Sub

Private Sub Cmdmid_Click()
Dim txt$, Middle$
txt$ = txtlft.Text
Middle$ = Mid(txt$, 3, 2)
MsgBox Middle$, vbOKOnly, "Middle String"
'The mid function is a little more complicated
'but not much.  3 is where the start of the middle
'string is, 2 is the lenth of the middle string
'so therefore since my string is "project1"it
'will return "oj".  "o" is the third letter
'of txt$ and the string that was returned was
'2 characters long(that is where the 2 came from.
End Sub

Private Sub CmdRevers_Click()
Dim txt$, back$
txt$ = txtlft.Text
back$ = StrReverse(txt$)
MsgBox back$, vbOKOnly, "Reverse"
'This function returns the whole string Reversed
End Sub

Private Sub Cmducase_Click()
Dim txt$, Upper$
txt$ = txtlft.Text
Upper$ = UCase(txt$)
MsgBox Upper$, vbOKOnly, "Uppercase"
'This function returns the whole string Uppercased
End Sub

Private Sub cmdlcase_Click()
Dim txt$, Lower$
txt$ = txtlft.Text
Lower$ = LCase(txt$)
MsgBox Lower$, vbOKOnly, "Lowercase"
'This function returns the whole string lowercased
End Sub

