VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form clipa 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Mostafa Live pictures"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   3
      Top             =   5835
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   1
      Min             =   7
      Max             =   8
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   960
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      Picture         =   "main.frx":0000
      ScaleHeight     =   1005
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   0
      Width           =   6045
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.microsoft.com"
      URL             =   "http://www.microsoft.com"
      RequestTimeout  =   420
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2175
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3836
      _Version        =   393217
      TextRTF         =   $"main.frx":19A6A
   End
End
Attribute VB_Name = "clipa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xx As Integer

Function IsConnected() As Boolean
On Error Resume Next

    Dim TRasCon(255) As RASCONN95
    Dim lg As Long
    Dim lpcon As Long
    Dim RetVal As Long
    Dim Tstatus As RASCONNSTATUS95
 
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
   
    RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)


    If RetVal <> 0 Then
        MsgBox "ERROR"
        Exit Function
    End If

   
    Tstatus.dwSize = 160
    RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)


    If Tstatus.RasConnState = &H2000 Then
        IsConnected = True
    Else
        IsConnected = False
    End If

End Function
Sub getpicture()
Dim b() As Byte
Dim xx
Dim strURL As String
Dim a As String
Screen.MousePointer = 11



RichTextBox1.Text = Inet1.OpenURL("http://www.geocities.com/ResearchTriangle/Campus/4598/clip1art.txt")
RichTextBox1.SaveFile App.Path & "\clip2.txt", 0
RichTextBox1.LoadFile App.Path & "\clip2.txt", 0
RichTextBox1.SaveFile App.Path & "\clip1.txt", 1



Dim freef As Integer
Dim nooo As Integer
freef = FreeFile
Open App.Path & "\clip1.txt" For Input As #freef
Do
Input #freef, a, strURL
b() = Inet1.OpenURL(strURL, icByteArray)
nooo = FreeFile
Open App.Path & "\temp.tmp" For Binary Access Write As #nooo
Put #nooo, , b()
Close #nooo
xx = xx + 1
ImageList1.ListImages.Add xx, , LoadPicture(App.Path & "\temp.tmp")
ListView1.ListItems.Add xx, , , xx, xx
Loop Until EOF(freef)
Close #freef
Screen.MousePointer = 0
End Sub
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.ScaleWidth
ListView1.Height = Me.Height - ListView1.Top - ProgressBar1.Height - 10
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
On Error Resume Next
If State >= 7 Then
ProgressBar1.Value = State
End If
End Sub

Private Sub ListView1_Click()
On Error GoTo er
imagemax = imagemax + 1
Load Form1.Image1(imagemax)
SavePicture ImageList1.ListImages.Item(ListView1.SelectedItem.Index).Picture, App.Path & "\imagelive" & ListView1.SelectedItem.Index & imagemax & ".bmp"
Form1.Image1(imagemax).Picture = LoadPicture(App.Path & "\imagelive" & ListView1.SelectedItem.Index & imagemax & ".bmp")
Form1.Image1(imagemax).ToolTipText = App.Path & "\imagelive" & ListView1.SelectedItem.Index & imagemax & ".bmp"
Form1.Image1(imagemax).Visible = True
Form1.Image1(imagemax).Left = 0
Form1.Image1(imagemax).Top = 0
Form1.Image1(imagemax).Stretch = True
Form1.Image1(imagemax).ZOrder 0
Dim str1 As String
Dim str2 As String
indexctrl = imagemax
changes = True
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If IsConnected = False Then
MsgBox "You must connect to internet before open Mostafa live pictures"
Unload Me
Exit Sub
Else
getpicture
Form_Resize
End If
Timer1.Enabled = False
End Sub
