VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Html code"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8190
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Image map"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add form"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Frame FRM 
      Caption         =   "Form"
      Height          =   2055
      Left            =   2400
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "Form1"
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "2"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "code.frx":0000
         Left            =   1080
         List            =   "code.frx":000A
         TabIndex        =   8
         Text            =   "POST"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Text            =   "action goes here"
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add it"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. of fields:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avalable tags"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write your html code on the textbox below:"
      Height          =   195
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tgg(100) As String

Private Sub Command1_Click()
On Error Resume Next
Dim i As Integer
Text1.SelStart = 0
Text1.Visible = False
For i = 0 To Len(Text1.Text)
Text1.SelStart = i
Text1.SelLength = 1
If Text1.SelText = """" Then
Text1.SelText = "'"
End If
Next i
Text1.Visible = True
Form1.Label3(ccindexx).Tag = Text1.Text
Unload Me
End Sub

Private Sub Command10_Click()
Dim i As Integer
Dim strg As String
FRM.Visible = False
    For i = 1 To Val(Text7.Text)
    strg = strg + "<p>enter your field label name<input type=" & """text""" & " size=" & """20""" & " name=""T" & i & """></p>"
Next i
Text1.SelText = "<form action=""" & Text8.Text & """ method=""" & Combo4.Text & """ name=""" & Text6.Text & """>" & strg & "<p><input type=" & """submit" & """ name=" & """B1""" & " value=" & """Submit""" & "><input type=" & """reset""" & "name=" & """B2""" & "value=" & """Reset""" & "></p>" & " </form>"

Text1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
FRM.Visible = True
End Sub

Private Sub Command4_Click()
On Error Resume Next
Form8.Show vbModal
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim taggg, codde
Open App.Path & "\tags.txt" For Input As #1
Do
Input #1, taggg
Input #1, codde
List1.AddItem taggg
tgg(List1.ListCount) = codde
Loop Until EOF(1)
Close #1
Form6.Text1.Text = Form1.Label3(ccindexx).Tag
End Sub

Private Sub List1_DblClick()
On Error Resume Next

Text1.SelText = tgg(List1.ListIndex + 1)
Text1.SetFocus
End Sub

