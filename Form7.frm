VERSION 5.00
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text editor"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2385
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   0
      Width           =   2380
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   60
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   12
         Top             =   3240
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   2400
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   0
         Left            =   60
         Picture         =   "Form7.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2760
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Text            =   "Arial"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   1695
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   600
         Width           =   2235
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   7
         Top             =   3240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   2
         Left            =   420
         Picture         =   "Form7.frx":03FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   1
         Left            =   780
         Picture         =   "Form7.frx":0644
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Left            =   1200
         Picture         =   "Form7.frx":088E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Height          =   375
         Left            =   1560
         Picture         =   "Form7.frx":0A5C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Height          =   375
         Left            =   1920
         Picture         =   "Form7.frx":0CAE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2760
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Opaque"
         Height          =   255
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Write your text here."
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   " Text editing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check4_Click()
On Error Resume Next
Form1.Label1(Text1.Tag).BackStyle = Check4.Value
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer
For i = 0 To Screen.FontCount - 1
Form7.Combo2.AddItem Screen.Fonts(i)
Next i
For i = 1 To 120
Form7.Combo3.AddItem i
Next i
End Sub

Private Sub Picture5_Click()
On Error GoTo er
MDI.CommonDialog1.ShowColor
Picture5.BackColor = MDI.CommonDialog1.Color
Form1.Label1(Text1.Tag).BackColor = Picture5.BackColor
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub
Private Sub Picture4_Click()
On Error GoTo er
MDI.CommonDialog1.ShowColor
Picture4.BackColor = MDI.CommonDialog1.Color
Form1.Label1(Text1.Tag).ForeColor = Picture4.BackColor
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Form1.Label1(Text1.Tag).FontName = Combo2.Text
changes = True
End Sub

Private Sub Combo3_Click()
On Error Resume Next
Form1.Label1(Text1.Tag).FontSize = Combo3.Text
changes = True
End Sub

Private Sub Combo4_Click()
On Error Resume Next
Form1.Shape1(Picture7.Tag).BorderStyle = Combo4.ListIndex
End Sub
Private Sub Check1_Click()
On Error Resume Next
Form1.Label1(Text1.Tag).FontBold = (Check1.Value <> 0)
End Sub

Private Sub Check2_Click()
On Error Resume Next
Form1.Label1(Text1.Tag).FontItalic = (Check2.Value <> 0)
End Sub

Private Sub Check3_Click()
On Error Resume Next
Form1.Label1(Text1.Tag).FontUnderline = (Check3.Value <> 0)
changes = True
End Sub
Private Sub Option1_Click(Index As Integer)
On Error Resume Next
For i = 0 To 2
If Option1(i).Value = True Then
 Form1.Label1(Text1.Tag).Alignment = i
End If
Next i
End Sub

Private Sub Text1_Change()
On Error Resume Next
Form1.Label1(Text1.Tag).Caption = Text1.Text
changes = True
If Text1.ForeColor = vbWhite Then
Text1.BackColor = vblack
Else
Text1.BackColor = vbWhite
End If
End Sub
