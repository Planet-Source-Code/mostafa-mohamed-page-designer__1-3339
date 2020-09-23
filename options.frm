VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Link"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5055
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image options"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er
MDI.CommonDialog1.Filter = "Gif images(*.gif)|*.gif|Jpeg inmages(*.jpg)|*.jpg|Windows bitmap(*.bmp)|*.bmp|Windows Metafile(*.wmf)|*.wmf|Icons(*.ico)|*.ico|Cursers(*.cur)|*.cur|"
MDI.CommonDialog1.ShowOpen
Form1.Image1(ccindexx).Stretch = False
Form1.Image1(ccindexx).Picture = LoadPicture(MDI.CommonDialog1.filename)
Form1.Image1(ccindexx).Stretch = True
Form1.Image1(ccindexx).ToolTipText = MDI.CommonDialog1.filename
Text1.Text = MDI.CommonDialog1.filename
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Me.Tag = "image" Then
Form1.Image1(indexctrl).ToolTipText = Text1.Text
Form1.Image1(indexctrl).Tag = Text2.Text
Unload Me
Else
Form1.Label1(indexctrl).Tag = Text2.Text
Unload Me
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Me.Tag = "image" Then
Frame1.Enabled = True
Else
Frame1.Enabled = False
End If
End Sub

