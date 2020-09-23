VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Page setup"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5985
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "..."
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "..."
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00000080&
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link color:"
      Height          =   195
      Left            =   2400
      TabIndex        =   26
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background picture:"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   1440
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4680
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background sound: "
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text color:- "
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Visited link:-"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Background color:-"
      Height          =   195
      Left            =   2400
      TabIndex        =   19
      Top             =   2880
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Title:-"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   390
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Page height:-"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   1000
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Page width:-"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   660
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filename:-"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   160
      Width           =   720
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Text2.Text = Val(Text2.Text) - 100
End Sub

Private Sub Command10_Click()
On Error GoTo er
MDI.CommonDialog1.Filter = "Gif images(*.gif)|*.gif|Jpeg inmages(*.jpg)|*.jpg|Windows bitmap(*.bmp)|*.bmp|Windows Metafile(*.wmf)|*.wmf|Icons(*.ico)|*.ico|Cursers(*.cur)|*.cur|"
MDI.CommonDialog1.ShowOpen
Text6.Text = MDI.CommonDialog1.filename
MDI.Picture11.Tag = MDI.CommonDialog1.filename
MDI.Picture11.Picture = LoadPicture(MDI.CommonDialog1.filename)
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Command11_Click()
On Error GoTo er
MDI.CommonDialog1.ShowColor
Command11.BackColor = MDI.CommonDialog1.Color
lclr = Command11.BackColor
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Text2.Text = Val(Text2.Text) + 100
End Sub

Private Sub Command3_Click()
On Error Resume Next
Text3.Text = Val(Text3.Text) + 100
End Sub

Private Sub Command4_Click()
On Error Resume Next
Text3.Text = Val(Text3.Text) - 100
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
On Error GoTo er
MDI.CommonDialog1.ShowColor
Command6.BackColor = MDI.CommonDialog1.Color
bgclr = Command6.BackColor
Form1.BackColor = Command6.BackColor
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Command7_Click()
On Error GoTo er
MDI.CommonDialog1.ShowColor
Command7.BackColor = MDI.CommonDialog1.Color
vclr = Command7.BackColor
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Command8_Click()
On Error GoTo er
MDI.CommonDialog1.ShowColor
Command8.BackColor = MDI.CommonDialog1.Color
tclr = Command8.BackColor
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Command9_Click()
On Error GoTo er
MDI.CommonDialog1.Filter = "Wav files(*.wav)|*.wav|Midi(*.mid)|*.mid|Aiff sound(*.aif;*.aifc;*.aiff)|*.aif;*.aifc;*.aiff|AU sound(*.au;*.snd)|*.au;*.snd|"
MDI.CommonDialog1.ShowOpen
Text5.Text = MDI.CommonDialog1.filename
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = curfile
Text2.Text = Form1.Width
Text3.Text = Form1.Height
Text4.Text = Form1.Tag
Text5.Text = bgsound
Text6.Text = MDI.Picture11.Tag
Command6.BackColor = bgclr
Command7.BackColor = vclr
Command8.BackColor = tclr
Command11.BackColor = lclr
End Sub

Private Sub Text2_Change()

 Form1.Width = Text2.Text
End Sub

Private Sub Text3_Change()
 Form1.Height = Text3.Text
End Sub

Private Sub Text4_Change()
On Error Resume Next


 Form1.Tag = Text4.Text
End Sub

Private Sub Text5_Change()
On Error Resume Next
bgsound = Text5.Text
End Sub
