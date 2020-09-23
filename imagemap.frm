VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New image map"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse image"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      Left            =   4920
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4320
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   0
      Top             =   840
      Width           =   4935
      Begin VB.PictureBox picview 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page url:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim start As Boolean
Dim selectr As Boolean
Dim x1 As Single
Dim y1 As Single
Dim x2 As Single
Dim y2 As Single
Dim filen As String
Dim html As String
Dim fcode As String

Private Sub Command1_Click()
On Error Resume Next
fcode = "<MAP NAME='map'>" & html & "<IMG SRC='" & filen & "' USEMAP='#map'></MAP>"
Form6.Text1.SelText = fcode
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo er
MDI.CommonDialog1.Filter = "Gif images(*.gif)|*.gif|Jpeg inmages(*.jpg)|*.jpg|Windows bitmap(*.bmp)|*.bmp|Windows Metafile(*.wmf)|*.wmf|Icons(*.ico)|*.ico|Cursers(*.cur)|*.cur|"
MDI.CommonDialog1.ShowOpen
picview.Picture = LoadPicture(MDI.CommonDialog1.filename)

filen = MDI.CommonDialog1.filename
If picview.Width > Picture1.ScaleWidth Then
HScroll1.Visible = True
HScroll1.Max = picview.Width - Picture1.ScaleWidth
Else
HScroll1.Visible = False
End If
If picview.Height > Picture1.ScaleHeight Then
VScroll1.Visible = True
VScroll1.Max = picview.Height - Picture1.ScaleHeight
Else
VScroll1.Visible = False
End If

Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Command3_Click()

On Error Resume Next
If Text1.Text <> "" Then
html = html + "<AREA SHAPE='RECT' COORDS='" & x1 & "," & y1 & "," & x2 & "," & y2 & "' HREF='" & Text1.Text & "'>"
Command3.Visible = False
Text1.Text = ""
Else
MsgBox "You must enter the page url"
End If
End Sub

Private Sub Form_Load()

start = True
selectr = False
On Error GoTo er
MDI.CommonDialog1.Filter = "Gif images(*.gif)|*.gif|Jpeg inmages(*.jpg)|*.jpg|Windows bitmap(*.bmp)|*.bmp|Windows Metafile(*.wmf)|*.wmf|Icons(*.ico)|*.ico|Cursers(*.cur)|*.cur|"
MDI.CommonDialog1.ShowOpen
picview.Picture = LoadPicture(MDI.CommonDialog1.filename)

filen = MDI.CommonDialog1.filename
If picview.Width > Picture1.ScaleWidth Then
HScroll1.Visible = True
HScroll1.Max = picview.Width - Picture1.ScaleWidth
Else
HScroll1.Visible = False
End If
If picview.Height > Picture1.ScaleHeight Then
VScroll1.Visible = True
VScroll1.Max = picview.Height - Picture1.ScaleHeight
Else
VScroll1.Visible = False
End If

Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub HScroll1_Change()
HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
On Error Resume Next
picview.Left = -HScroll1.Value
End Sub

Private Sub picview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If start = True Then
Command3.Visible = False
x1 = X
y1 = Y
start = False
selectr = True
End If

End Sub

Private Sub picview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If selectr = True Then
picview.Cls
Rectangle picview.hdc, x1, y1, X, Y
End If
End Sub

Private Sub picview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If selectr = True Then
Command3.Visible = True
picview.Cls
Rectangle picview.hdc, x1, y1, X, Y
y2 = Y
x2 = X
start = True
selectr = False
End If
End Sub

Private Sub VScroll1_Change()
VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
On Error Resume Next
picview.Top = -VScroll1.Value
End Sub
