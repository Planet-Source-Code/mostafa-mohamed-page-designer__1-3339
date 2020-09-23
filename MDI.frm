VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Mostafa-Page designer"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   3  'Align Left
      Height          =   6225
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   10980
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "image"
            Object.ToolTipText     =   "Image"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "text"
            Object.ToolTipText     =   "Text"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tag"
            Object.ToolTipText     =   "Html tag"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "line"
            Object.ToolTipText     =   "Line"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mouse"
            Object.ToolTipText     =   "Mouse"
            ImageIndex      =   7
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin VB.PictureBox Picture11 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   615
         TabIndex        =   6
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3720
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":0F5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":1BB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":2806
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":345A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":3776
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":38DA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3600
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":3BF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":3D0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":3E1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":3F32
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":4046
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDI.frx":415A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture13 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   0
      ScaleHeight     =   1725
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   6585
      Width           =   11880
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   7080
         TabIndex        =   7
         Text            =   "http://www.geocities.com/ResearchTriangle/Campus/4598/pd.html"
         ToolTipText     =   "Address"
         Top             =   0
         Width           =   3615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Refresh"
         Height          =   315
         Left            =   5520
         TabIndex        =   12
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Next"
         Height          =   315
         Left            =   5040
         TabIndex        =   9
         ToolTipText     =   "Go next"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Home"
         Height          =   315
         Left            =   4440
         TabIndex        =   11
         ToolTipText     =   "Go home"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Back"
         Height          =   315
         Left            =   3960
         TabIndex        =   8
         ToolTipText     =   "Go back"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   0
         Width           =   2415
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Update view"
         Height          =   315
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Update the view of your design on the browser"
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "[---]"
         Height          =   315
         Left            =   0
         TabIndex        =   10
         ToolTipText     =   "Large or dislarge the browser window"
         Top             =   0
         Width           =   375
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   6495
         Left            =   0
         TabIndex        =   1
         Top             =   315
         Width           =   3735
         ExtentX         =   6588
         ExtentY         =   11456
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6360
         TabIndex        =   13
         Top             =   45
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnunew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnysaveas 
         Caption         =   "Save as"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuprintsetup 
         Caption         =   "Print setup"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnucut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnupagesetup 
         Caption         =   "Page setup"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuas 
         Caption         =   "Actual size"
         Shortcut        =   {F9}
      End
      Begin VB.Menu nbup 
         Caption         =   "Zoom up"
      End
      Begin VB.Menu mnudwn 
         Caption         =   "Zoom down"
      End
   End
   Begin VB.Menu mnuinsert 
      Caption         =   "Insert"
      Begin VB.Menu mnuimage 
         Caption         =   "Image"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnutext 
         Caption         =   "Text"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuline 
         Caption         =   "Line"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuhtml 
         Caption         =   "Html code"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnulive 
         Caption         =   "Live pictures"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu mnuexp 
      Caption         =   "Export html"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnucontent 
         Caption         =   "Content"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnubuy 
         Caption         =   "Buy it"
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         HelpContextID   =   1
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancela As Boolean
Dim newnum As Integer
Private Const SC_CLOSE As Long = &HF060&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86


Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
    End Type


Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hwnd As Long, ByVal bRevert As Long) As Long


Private Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long


Private Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long


Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long


Private Declare Function IsWindow Lib "user32" _
    (ByVal hwnd As Long) As Long

Dim larged As Boolean
Public Function EnableCloseButton(ByVal hwnd As Long, Enable As Boolean) As Integer
    Const xSC_CLOSE As Long = -10
    ' Check that the window handle passed is valid
    EnableCloseButton = -1
    If IsWindow(hwnd) = 0 Then Exit Function
    ' Retrieve a handle to the window's system menu
    Dim hMenu As Long
    hMenu = GetSystemMenu(hwnd, 0)
    ' Retrieve the menu item information for the close menu item/butt
    '     on
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE


    If Enable Then
        MII.wID = xSC_CLOSE
    Else
        MII.wID = SC_CLOSE
    End If

    EnableCloseButton = -0
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    ' Switch the ID of the menu item so that VB can not undo the acti
    '     on itself
    Dim lngMenuID As Long
    lngMenuID = MII.wID


    If Enable Then
        MII.wID = SC_CLOSE
    Else
        MII.wID = xSC_CLOSE
    End If

    MII.fMask = MIIM_ID
    EnableCloseButton = -2
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function
    ' Set the enabled / disabled state of the menu item


    If Enable Then
        MII.fState = (MII.fState Or MFS_GRAYED)
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If

    MII.fMask = MIIM_STATE
    EnableCloseButton = -3
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    ' Activate the non-client area of the window to update the titleb
    '     ar, and
    ' draw the close button in its new state.
    SendMessage hwnd, WM_NCACTIVATE, True, 0
    EnableCloseButton = 0
End Function

Sub getactualsize()
If zoo = 1 Then
Form1.zoomdown 2
GoTo dozoo
End If
If zoo = -1 Then
Form1.zoomup 2
GoTo dozoo
End If
dozoo:
zoo = 0
End Sub
Sub lneput(X As Single, Y As Single)
On Error GoTo er
linemax = linemax + 1
Load Form1.Label2(linemax)
Form1.Label2(linemax).Visible = True
Form1.Label2(linemax).Left = X
Form1.Label2(linemax).Top = Y
Form1.Label2(linemax).BackColor = tclr
Form1.Label2(linemax).ZOrder 0
changes = True
Exit Sub
er:
MsgBox Err.Description
End Sub
Sub hmlput(X As Single, Y As Single)
On Error GoTo er
cmax = cmax + 1
Load Form1.Label3(cmax)
Form1.Label3(cmax).Visible = True
Form1.Label3(cmax).Left = X
Form1.Label3(cmax).Top = Y
Form1.Label3(cmax).ZOrder 0
changes = True
Exit Sub
er:
MsgBox Err.Description
End Sub
Sub txtput(X As Single, Y As Single)
On Error GoTo er

textmax = textmax + 1
Load Form1.Label1(textmax)
Form1.Label1(textmax).Visible = True
Form1.Label1(textmax).Left = X
Form1.Label1(textmax).Top = Y
Form1.Label1(textmax).ForeColor = tclr
Form1.Label1(textmax).ZOrder 0
changes = True
Exit Sub
er:
MsgBox Err.Description
End Sub
Sub imgput(X As Single, Y As Single)
On Error GoTo er
imagemax = imagemax + 1
Load Form1.Image1(imagemax)
Form1.Image1(imagemax).ToolTipText = CommonDialog1.filename
Form1.Image1(imagemax).Visible = True
Form1.Image1(imagemax).Left = X
Form1.Image1(imagemax).Top = Y
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
Sub paste()
On Error Resume Next
Dim no As Integer
Dim i As Integer

no = FreeFile
Open App.Path & "\clip.txt" For Input As #no

Input #no, a
Input #no, b
Input #no, c
Input #no, d
Input #no, e
Input #no, f
Input #no, g
Input #no, h
Input #no, i
Input #no, j
Input #no, k
Input #no, l
Input #no, m
Input #no, n
Input #no, o
Input #no, p
Input #no, Q
Input #no, r

If a = "Label" Then
If r = "Text" Then
textmax = textmax + 1
Load Form1.Label1(textmax)
Form1.Label1(textmax).Caption = b
Form1.Label1(textmax).Width = c
Form1.Label1(textmax).Height = d
Form1.Label1(textmax).Left = e + 500
Form1.Label1(textmax).Top = f + 500
Form1.Label1(textmax).BackStyle = i
Form1.Label1(textmax).BackColor = g
Form1.Label1(textmax).ForeColor = h
Form1.Label1(textmax).FontName = j
Form1.Label1(textmax).FontSize = k
Form1.Label1(textmax).FontBold = l
Form1.Label1(textmax).FontUnderline = m
Form1.Label1(textmax).FontItalic = n
Form1.Label1(textmax).Alignment = o
Form1.Label1(textmax).Tag = Q
Form1.Label1(textmax).ToolTipText = r
Form1.Label1(textmax).ZOrder 0
tcde(textmax) = p
Form1.Label1(textmax).Visible = True
textmax = textmax + 1
End If
If r = "Line" Then
linemax = linemax + 1
Load Form1.Label2(linemax)
Form1.Label2(linemax).Caption = b
Form1.Label2(linemax).Width = c
Form1.Label2(linemax).Height = d
Form1.Label2(linemax).Left = e + 500
Form1.Label2(linemax).Top = f + 500
Form1.Label2(linemax).BackStyle = i
Form1.Label2(linemax).BackColor = g
Form1.Label2(linemax).ForeColor = h
Form1.Label2(linemax).FontName = j
Form1.Label2(linemax).FontSize = k
Form1.Label2(linemax).FontBold = l
Form1.Label2(linemax).FontUnderline = m
Form1.Label2(linemax).FontItalic = n
Form1.Label2(linemax).Alignment = o
Form1.Label2(linemax).ToolTipText = r
Form1.Label2(linemax).Tag = Q
Form1.Label2(linemax).ZOrder 0
Form1.Label2(linemax).Visible = True
linemax = linemax + 1
End If
If r = "Html code" Then
cmax = cmax + 1
Load Form1.Label3(cmax)
Form1.Label3(cmax).Caption = b
Form1.Label3(cmax).Width = c
Form1.Label3(cmax).Height = d
Form1.Label3(cmax).Left = e + 500
Form1.Label3(cmax).Top = f + 500
Form1.Label3(cmax).BackStyle = i
Form1.Label3(cmax).BackColor = g
Form1.Label3(cmax).ForeColor = h
Form1.Label3(cmax).FontName = j
Form1.Label3(cmax).FontSize = k
Form1.Label3(cmax).FontBold = l
Form1.Label3(cmax).FontUnderline = m
Form1.Label3(cmax).FontItalic = n
Form1.Label3(cmax).Alignment = o
Form1.Label3(cmax).ToolTipText = r
Form1.Label3(cmax).Tag = Q
Form1.Label3(cmax).ZOrder 0
Form1.Label3(cmax).Visible = True
cmax = cmax + 1
End If
End If

If a = "Shape" Then
shapemax = shapemax + 1
Load Form1.Shape1(shapemax)
Form1.Shape1(shapemax).Shape = b
Form1.Shape1(shapemax).Width = c
Form1.Shape1(shapemax).Height = d
Form1.Shape1(shapemax).Left = e + 500
Form1.Shape1(shapemax).Top = f + 500
Form1.Shape1(shapemax).BackStyle = i
Form1.Shape1(shapemax).BackColor = g
Form1.Shape1(shapemax).BorderColor = h
Form1.Shape1(shapemax).BorderStyle = j
Form1.Shape1(shapemax).BorderWidth = k
Form1.Shape1(shapemax).Visible = True
shapemax = shapemax + 1
End If

If a = "Image" Then
imagemax = imagemax + 1
Load Form1.Image1(imagemax)
Form1.Image1(imagemax).Width = b
Form1.Image1(imagemax).Height = c
Form1.Image1(imagemax).Left = d + 500
Form1.Image1(imagemax).Top = e + 500
Form1.Image1(imagemax).ToolTipText = f
Form1.Image1(imagemax).Stretch = True
Form1.Image1(imagemax).Visible = True
Form1.Image1(imagemax).Picture = LoadPicture(f)
Form1.Image1(imagemax).Tag = Q
Form1.Image1(imagemax).ZOrder 0
icde(imagemax) = p
imagemax = imagemax + 1
End If



Close #no
End Sub
Sub copy()
On Error Resume Next
Dim no As Integer
Dim i As Integer
no = FreeFile


Open App.Path & "\clip.txt" For Output As #no
If ctrltype = "text" Then
Write #no, "Label"
Write #no, Form1.Label1(ccindexx).Caption
Write #no, Form1.Label1(ccindexx).Width
Write #no, Form1.Label1(ccindexx).Height
Write #no, Form1.Label1(ccindexx).Left
Write #no, Form1.Label1(ccindexx).Top
Write #no, Form1.Label1(ccindexx).BackColor
Write #no, Form1.Label1(ccindexx).ForeColor
Write #no, Form1.Label1(ccindexx).BackStyle
Write #no, Form1.Label1(ccindexx).FontName
Write #no, Form1.Label1(ccindexx).FontSize
Write #no, Form1.Label1(ccindexx).FontBold
Write #no, Form1.Label1(ccindexx).FontUnderline
Write #no, Form1.Label1(ccindexx).FontItalic
Write #no, Form1.Label1(ccindexx).Alignment
Write #no, tcde(Form1.Label1(ccindexx).Index)
Write #no, Form1.Label1(ccindexx).Tag
Write #no, Form1.Label1(ccindexx).ToolTipText
End If
If ctrltype = "html" Then
Write #no, "Label"
Write #no, Form1.Label3(ccindexx).Caption
Write #no, Form1.Label3(ccindexx).Width
Write #no, Form1.Label3(ccindexx).Height
Write #no, Form1.Label3(ccindexx).Left
Write #no, Form1.Label3(ccindexx).Top
Write #no, Form1.Label3(ccindexx).BackColor
Write #no, Form1.Label3(ccindexx).ForeColor
Write #no, Form1.Label3(ccindexx).BackStyle
Write #no, Form1.Label3(ccindexx).FontName
Write #no, Form1.Label3(ccindexx).FontSize
Write #no, Form1.Label3(ccindexx).FontBold
Write #no, Form1.Label3(ccindexx).FontUnderline
Write #no, Form1.Label3(ccindexx).FontItalic
Write #no, Form1.Label3(ccindexx).Alignment
Write #no, tcde(Form1.Label3(ccindexx).Index)
Write #no, Form1.Label3(ccindexx).Tag
Write #no, Form1.Label3(ccindexx).ToolTipText
End If
If ctrltype = "line" Then
Write #no, "Label"
Write #no, Form1.Label2(ccindexx).Caption
Write #no, Form1.Label2(ccindexx).Width
Write #no, Form1.Label2(ccindexx).Height
Write #no, Form1.Label2(ccindexx).Left
Write #no, Form1.Label2(ccindexx).Top
Write #no, Form1.Label2(ccindexx).BackColor
Write #no, Form1.Label2(ccindexx).ForeColor
Write #no, Form1.Label2(ccindexx).BackStyle
Write #no, Form1.Label2(ccindexx).FontName
Write #no, Form1.Label2(ccindexx).FontSize
Write #no, Form1.Label2(ccindexx).FontBold
Write #no, Form1.Label2(ccindexx).FontUnderline
Write #no, Form1.Label2(ccindexx).FontItalic
Write #no, Form1.Label2(ccindexx).Alignment
Write #no, tcde(Form1.Label2(ccindexx).Index)
Write #no, Form1.Label2(ccindexx).Tag
Write #no, Form1.Label2(ccindexx).ToolTipText
End If


If ctrltype = "image" Then
Write #no, "Image"
Write #no, Form1.Image1(ccindexx).Width
Write #no, Form1.Image1(ccindexx).Height
Write #no, Form1.Image1(ccindexx).Left
Write #no, Form1.Image1(ccindexx).Top
Write #no, Form1.Image1(ccindexx).ToolTipText
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, "none"
Write #no, icde(Form1.Controls(i).Index)
Write #no, Form1.Controls(i).Tag
End If
Close #no
End Sub
Sub delete()
On Error Resume Next
Form1.ShowHandles False
Form1.DragEnd
If ctrltype = "text" Then
Unload Form1.Label1(ccindexx)
End If
If ctrltype = "line" Then
Unload Form1.Label2(ccindexx)
End If
If ctrltype = "html" Then
Unload Form1.Label3(ccindexx)
End If
If ctrltype = "image" Then
Unload Form1.Image1(ccindexx)
End If
End Sub
Sub newpage()
On Error Resume Next

Dim res As VbMsgBoxResult
If dirty = True Then
res = MsgBox("Do you want to save changes to current page?", vbYesNoCancel)
If res = vbYes Then
saveas
getactualsize

curfile = ""
Unload Form1
Form1.Show
dirty = False
End If
If res = vbNo Then
getactualsize
curfile = ""
Unload Form1
Form1.Show
dirty = False
End If
If res = vbCancel Then
Exit Sub
End If
Else
getactualsize
curfile = ""
Unload Form1
Form1.Show
dirty = False
End If
End Sub
Sub exitit()
On Error Resume Next

Dim res As VbMsgBoxResult
If dirty = True Then
res = MsgBox("Do you want to save changes to current page?", vbYesNoCancel)
If res = vbYes Then
saveas
getactualsize

curfile = ""
End
dirty = False
End If
If res = vbNo Then
getactualsize
curfile = ""
End
dirty = False
End If
If res = vbCancel Then

Exit Sub
End If
Else
End

End If
End Sub

Sub saveas()
On Error GoTo errr:

If curfile = "" Then
With MDI.CommonDialog1
getactualsize
.Filter = "Page designer files (*.mpf)|*.mpf|All files (*.*)|*.|"
.ShowSave
Savepage .filename
curfile = .filename
dirty = False
End With
Exit Sub
Else
getactualsize
Savepage curfile
dirty = False
Exit Sub
End If
Exit Sub
errr:
If Err.Number = 32755 Then
Exit Sub
Else
MsgBox Err.Description
End If
End Sub
Sub openpge()
On Error GoTo errr
Dim res As VbMsgBoxResult
If dirty = True Then
res = MsgBox("Do you want to save changes to current page?", vbYesNoCancel)
If res = vbYes Then
saveas
getactualsize
With MDI.CommonDialog1

.Filter = "Page designer files (*.mpf)|*.mpf|All files (*.*)|*.|"
.ShowOpen
curfile = .filename
Unload Form1
Form1.Show
Openpage .filename
dirty = False
End With
curfile = ""
End If
If res = vbNo Then
With MDI.CommonDialog1
getactualsize
.Filter = "Page designer files (*.mpf)|*.mpf|All files (*.*)|*.|"
.ShowOpen
curfile = .filename
Unload Form1
Form1.Show
Openpage .filename
dirty = False
End With
End If
If res = vbCancel Then
Exit Sub
End If

Else
With MDI.CommonDialog1
getactualsize
.Filter = "Page designer files (*.mpf)|*.mpf|All files (*.*)|*.|"
.ShowOpen
curfile = .filename
Unload Form1
Form1.Show
Openpage .filename
End With
dirty = False
End If
Exit Sub
errr:
If Err.Number = 32755 Then
Exit Sub
Else
MsgBox Err.Description
curfile = ""
End If
End Sub

Sub saveit(filename As String)
On Error GoTo er
Savepage filename
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub








Private Sub BtnGraphic6_Click()
End
End Sub

Private Sub Check1_Click()
Form1.Label1(Text1.Tag).FontBold = (Check1.Value <> 0)
End Sub

Private Sub Check2_Click()
Form1.Label1(Text1.Tag).FontItalic = (Check2.Value <> 0)
End Sub

Private Sub Check3_Click()
Form1.Label1(Text1.Tag).FontUnderline = (Check3.Value <> 0)
changes = True
End Sub

Private Sub Check4_Click()
On Error Resume Next
Form1.Label1(Text1.Tag).BackStyle = Check4.Value
End Sub

Private Sub Check5_Click()
Form1.Shape1(Picture7.Tag).BackStyle = Check5.Value
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

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
WebBrowser1.Navigate Combo1.Text
Combo1.AddItem Combo1.Text
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub

Private Sub Command2_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Picture13.Height = 1725 Then
Picture13.Height = 4725
Exit Sub
Else
Picture13.Height = 1725
Exit Sub
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
WebBrowser1.Navigate "http://www.geocities.com/ResearchTriangle/Campus/4598/pd.html"
End Sub

Private Sub Command5_Click()
On Error Resume Next
WebBrowser1.Refresh
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command7_Click()
On Error Resume Next
getactualsize
savehtml App.Path & "\tmp.html"
WebBrowser1.Navigate App.Path & "\tmp.html"
End Sub

Private Sub Command8_Click()
Form1.Width = Form1.Width - 1000

End Sub

Private Sub Command9_Click()
paste
End Sub

Private Sub Label1_Click()
Picture6.Visible = False
End Sub

Private Sub Label11_Click()
newpage
End Sub

Private Sub Label16_Click(Index As Integer)
On Error GoTo er
shapemax = shapemax + 1
Load Form1.Shape1(shapemax)
Form1.Shape1(shapemax).Visible = True
Form1.Shape1(shapemax).Left = 0
Form1.Shape1(shapemax).Top = 0
Form1.Shape1(shapemax).Shape = Index
Picture6.Visible = False
changes = True
Exit Sub
er:
MsgBox Err.Description
End Sub

Private Sub Label17_Click()
On Error GoTo er
CommonDialog1.Filter = "Windows bitmap(*.bmp)|*.bmp|Windows Metafile(*.wmf)|*.wmf|Gif images(*.gif)|*.gif|Jpeg inmages(*.jpg)|*.jpg|Icons(*.ico)|*.ico|Cursers(*.cur)|*.cur|"
CommonDialog1.ShowOpen
Picture11.Picture = LoadPicture(CommonDialog1.filename)
Picture11.Tag = CommonDialog1.filename
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Label18_Click()
openpge
changes = False
End Sub

Private Sub Label19_Click()
saveas
End Sub

Private Sub Label20_Click()
With MDI.CommonDialog1
.Filter = "Page designer files (*.mpf)|*.mpf|All files (*.*)|*.|"
.ShowSave
Savepage .filename
curfile = .filename
End With
End Sub

Private Sub Label21_Click()
Form3.Show vbModal
End Sub

Private Sub Label3_Click()
Picture6.Visible = False
End Sub

Private Sub Label5_Click(Index As Integer)
Picture6.Visible = False
End Sub

Private Sub MDIForm_Load()
On Error Resume Next
Dim X As Integer
curfile = ""
Form1.Show
curfile = ""
dirty = False
Combo1.Text = "http://www.geocities.com/ResearchTriangle/Campus/4598/pd.html"

X = EnableCloseButton(Me.hwnd, False)

Open "c:\no.fle" For Input As #2
Input #2, newnum
If newnum >= 30 Then
MsgBox "Trial version expired"
Kill "c:\fond.fle"
End
Exit Sub
End If
Close #2

Open "c:\date.fle" For Input As #1
Input #1, d
l = Format(d, "dd")
X = Format(Date, "dd")

If l <> X Then
Open "c:\no.fle" For Output As #3
Write #3, newnum + 1
Close #3
End If
Close #1
Open "c:\date.fle" For Output As #4
Write #4, Date
Close #4

Open "c:\fond.fle" For Input As #5
Input #5, c
Close #5
If c = "fond" Then

Else
Open "c:\fond.fle" For Input As #6
Write #6, "fond"
Close #6
Open "c:\first.fle" For Input As #7
Input #7, f
Close #7
If f = "new" Then
MsgBox "program already installed and expired"
Kill "c:\fond.fle"
End
Exit Sub
Else
Open "c:\fond.fle" For Output As #9
Write #9, "fond"
Close #9
Open "c:\first.fle" For Output As #8
Write #8, "new"
Close #8

End If
End If
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
Form1.Left = ((MDI.Width - Toolbar2.Width) - Form1.Width) / 2
Form1.Top = (MDI.Height - Form1.Height - MDI.Picture13.Height) / 2
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
'exitit
End Sub

Private Sub mnuabout_Click()
On Error Resume Next
frmAbout.Show vbModal
End Sub

Private Sub mnuas_Click()
getactualsize
End Sub

Private Sub mnubuy_Click()
On Error Resume Next
WebBrowser1.Navigate "http://www.geocities.com/ResearchTriangle/Campus/4598/pdbuy.html"
End Sub

Private Sub mnucontent_Click()
On Error Resume Next
    Dim Scr_hDC As Long
    Dim startdoc As Long
    Dim l, txtt
    Scr_hDC = GetDesktopWindow()
    l = InStr(1, App.Path, "\", 1)
    txtt = Mid(App.Path, 1, l)
    startdoc = ShellExecute(Scr_hDC, "Open", App.Path & "\PDhelp.chm", "", txtt, SW_SHOWNORMAL)
       
        If Err Then
           MsgBox Err.Description
        End If
End Sub

Private Sub mnucopy_Click()
copy
End Sub

Private Sub mnucut_Click()
copy
delete
End Sub

Private Sub mnudelete_Click()
delete
End Sub

Private Sub mnudwn_Click()
If zoo > -1 Then
Form1.zoomdown 2
End If
End Sub

Private Sub mnuexit_Click()
exitit
End Sub

Private Sub mnuexp_Click()
On Error GoTo errr:
With MDI.CommonDialog1
.Filter = "Html files (*.html)|*.html|Htm files (*.htm)|*.htm|All files (*.*)|*.|"
.ShowSave
getactualsize
savehtml .filename
End With
Exit Sub
errr:
If Err.Number = 32755 Then
Exit Sub
Else
MsgBox Err.Description
End If
End Sub

Private Sub mnuhtled_Click()
On Error Resume Next
Shell App.Path & "\web page editor.exe", vbNormalFocus
End Sub

Private Sub mnuhtml_Click()
addctrl = "tag"
End Sub

Private Sub mnuimage_Click()
addctrl = "image"
End Sub

Private Sub mnuline_Click()
addctrl = "line"
End Sub

Private Sub mnulive_Click()
clipa.Show
End Sub

Private Sub mnunew_Click()
newpage
End Sub

Private Sub mnuopen_Click()
openpge
End Sub

Private Sub mnupagesetup_Click()
Form3.Show vbModal
End Sub

Private Sub mnupaste_Click()
paste
End Sub

Private Sub mnuprint_Click()
On Error Resume Next
Form1.ShowHandles False
getactualsize
Form1.PrintForm
End Sub

Private Sub mnuprintsetup_Click()
On Error GoTo errr:
With MDI.CommonDialog1
.ShowPrinter
End With
Exit Sub
errr:
If Err.Number = 32755 Then
Exit Sub
Else
MsgBox Err.Description
End If
End Sub

Private Sub mnusave_Click()
saveas
End Sub

Private Sub mnutext_Click()
addctrl = "text"
End Sub

Private Sub mnysaveas_Click()
With MDI.CommonDialog1
.Filter = "Page designer files (*.mpf)|*.mpf|All files (*.*)|*.|"
.ShowSave
Savepage .filename
curfile = .filename
End With
End Sub





Private Sub nbup_Click()
If zoo < 1 Then
Form1.zoomup 2
End If
End Sub

Private Sub Picture13_Resize()
On Error Resume Next
WebBrowser1.Height = Picture13.Height - WebBrowser1.Top
WebBrowser1.Width = Picture13.Width - WebBrowser1.Left
Combo1.Width = Picture13.Width - Combo1.Left
End Sub


























Private Sub UButton1_Click()

End Sub

Private Sub UButton1_GotMouse()
MyAgent.Stop
MyAgent.Speak "Click here to add a picture"

End Sub

Private Sub UButton2_Click()

End Sub

 Private Sub UButton2_GotMouse()
 MyAgent.Stop
MyAgent.Speak "Click here to add text"
End Sub

Private Sub UButton3_Click()

End Sub

Private Sub UButton3_GotMouse()
MyAgent.Stop
MyAgent.Speak "Click here to add shape"
End Sub











Private Sub Timer1_Timer()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "new"
newpage
Case "open"
openpge
Case "save"
saveas
Case "cut"
copy
delete
Case "copy"
copy
Case "paste"
paste
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "image"
addctrl = "image"
Case "text"
addctrl = "text"
Case "tag"
addctrl = "tag"
Case "line"
addctrl = "line"
Case "mouse"
addctrl = "mouse"
End Select
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
On Error Resume Next
Command8.Caption = Text
End Sub
