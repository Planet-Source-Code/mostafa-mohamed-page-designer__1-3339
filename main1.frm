VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10965
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   9210
   Begin VB.PictureBox picHandle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2520
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Html code  "
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "Html code"
      Top             =   1440
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   30
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Line"
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   0
      Left            =   1560
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   360
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add Text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Text"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Const NULL_BRUSH = 5
Private Const PS_SOLID = 0
Private Const R2_NOT = 6

Enum ControlState
    StateNothing = 0
    StateDragging
    StateSizing
End Enum

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New CRect
Private m_DragPoint As POINTAPI

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Sub zoomup(vle As Integer)
On Error Resume Next
Dim i As Integer

Form1.Width = Form1.Width * vle
Form1.Height = Form1.Height * vle
Form1.ShowHandles False
For i = 0 To Form1.Controls.Count - 1
Form1.Controls(i).Left = Form1.Controls(i).Left * vle
Form1.Controls(i).Top = Form1.Controls(i).Top * vle
Form1.Controls(i).Width = Form1.Controls(i).Width * vle
Form1.Controls(i).Height = Form1.Controls(i).Height * vle
Form1.Controls(i).FontSize = Form1.Controls(i).FontSize * vle
Next i
zoo = zoo + 1
End Sub
Sub zoomdown(vle As Integer)
On Error Resume Next
Dim i As Integer

Form1.Width = Form1.Width / vle
Form1.Height = Form1.Height / vle
Form1.ShowHandles False
For i = 0 To Form1.Controls.Count - 1
Form1.Controls(i).Left = Form1.Controls(i).Left / vle
Form1.Controls(i).Top = Form1.Controls(i).Top / vle
Form1.Controls(i).Width = Form1.Controls(i).Width / vle
Form1.Controls(i).Height = Form1.Controls(i).Height / vle
Form1.Controls(i).FontSize = Form1.Controls(i).FontSize / vle
Next i
zoo = zoo - 1
End Sub
Sub ShowHandles(Optional bShowHandles As Boolean = True)
    Dim i As Integer
    Dim xFudge As Long, yFudge As Long
    Dim nWidth As Long, nHeight As Long

    If bShowHandles And Not m_CurrCtl Is Nothing Then
        With m_DragRect
           
            nWidth = (picHandle(0).Width \ 2)
            nHeight = (picHandle(0).Height \ 2)
            xFudge = (0.5 * Screen.TwipsPerPixelX)
            yFudge = (0.5 * Screen.TwipsPerPixelY)
           
            picHandle(0).Move (.Left - nWidth) + xFudge, (.Top - nHeight) + yFudge
            
            picHandle(4).Move (.Left + .Width) - nWidth - xFudge, .Top + .Height - nHeight - yFudge
          
            picHandle(1).Move .Left + (.Width / 2) - nWidth, .Top - nHeight + yFudge
           
            picHandle(5).Move .Left + (.Width / 2) - nWidth, .Top + .Height - nHeight - yFudge
          
            picHandle(2).Move .Left + .Width - nWidth - xFudge, .Top - nHeight + yFudge
         
            picHandle(6).Move .Left - nWidth + xFudge, .Top + .Height - nHeight - yFudge
         
            picHandle(3).Move .Left + .Width - nWidth - xFudge, .Top + (.Height / 2) - nHeight
            
            picHandle(7).Move .Left - nWidth + xFudge, .Top + (.Height / 2) - nHeight
        End With
    End If
  
    For i = 0 To 7
        picHandle(i).Visible = bShowHandles
    Next i
End Sub
Private Sub DrawDragRect()
    Dim hPen As Long, hOldPen As Long
    Dim hBrush As Long, hOldBrush As Long
    Dim hScreenDC As Long, nDrawMode As Long

   
    hScreenDC = GetDC(0)
  
    hPen = CreatePen(PS_SOLID, 2, 0)
    hOldPen = SelectObject(hScreenDC, hPen)
    hBrush = GetStockObject(NULL_BRUSH)
    hOldBrush = SelectObject(hScreenDC, hBrush)
    nDrawMode = SetROP2(hScreenDC, R2_NOT)
  
    Rectangle hScreenDC, m_DragRect.Left, m_DragRect.Top, _
        m_DragRect.Right, m_DragRect.Bottom

    SetROP2 hScreenDC, nDrawMode
    SelectObject hScreenDC, hOldBrush
    SelectObject hScreenDC, hOldPen
    ReleaseDC 0, hScreenDC

    DeleteObject hPen
End Sub

Private Sub DragBegin(ctl As Control)
    Dim rc As RECT

 
    ShowHandles False
  
    Set m_CurrCtl = ctl
  
    GetCursorPos m_DragPoint
   
    m_DragRect.SetRectToCtrl m_CurrCtl
    m_DragRect.TwipsToScreen m_CurrCtl
   
    m_DragPoint.X = m_DragPoint.X - m_DragRect.Left
    m_DragPoint.Y = m_DragPoint.Y - m_DragRect.Top

    Refresh

    DrawDragRect
  
    m_DragState = StateDragging
  
    ReleaseCapture
    SetCapture hwnd

    GetWindowRect hwnd, rc
    ClipCursor rc
End Sub


 Sub DragEnd()
    Set m_CurrCtl = Nothing
    ShowHandles False
    m_DragState = StateNothing
End Sub
Private Sub DragInit()
    Dim i As Integer, xHandle As Single, yHandle As Single

   
    xHandle = 5 * Screen.TwipsPerPixelX
    yHandle = 5 * Screen.TwipsPerPixelY
   
    For i = 0 To 7
        If i <> 0 Then
            Load picHandle(i)
        End If
        picHandle(i).Width = xHandle
        picHandle(i).Height = yHandle

        picHandle(i).ZOrder
    Next i

    picHandle(0).MousePointer = vbSizeNWSE
    picHandle(1).MousePointer = vbSizeNS
    picHandle(2).MousePointer = vbSizeNESW
    picHandle(3).MousePointer = vbSizeWE
    picHandle(4).MousePointer = vbSizeNWSE
    picHandle(5).MousePointer = vbSizeNS
    picHandle(6).MousePointer = vbSizeNESW
    picHandle(7).MousePointer = vbSizeWE

    Set m_CurrCtl = Nothing
End Sub

Private Sub BtnGraphic1_Click()
End
End Sub

Private Sub doc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer

    If Button = vbLeftButton And m_bDesignMode Then
      
        For i = 0 To (Controls.Count - 1)
           
            If Not TypeOf Controls(i) Is Menu And Controls(i).Visible Then
                m_DragRect.SetRectToCtrl Controls(i)
                If m_DragRect.PtInRect(X, Y) Then
                    DragBegin Controls(i)
                    Exit Sub
                End If
            End If
        Next i
 
        Set m_CurrCtl = Nothing
   
        ShowHandles False
    End If
End Sub

Private Sub doc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nWidth As Single, nHeight As Single
    Dim pt As POINTAPI

    If m_DragState = StateDragging Then
   
        nWidth = m_DragRect.Right - m_DragRect.Left
        nHeight = m_DragRect.Bottom - m_DragRect.Top
       
        GetCursorPos pt
  
        DrawDragRect
   
        m_DragRect.Left = pt.X - m_DragPoint.X
        m_DragRect.Top = pt.Y - m_DragPoint.Y
        m_DragRect.Right = m_DragRect.Left + nWidth
        m_DragRect.Bottom = m_DragRect.Top + nHeight
    
        DrawDragRect
    ElseIf m_DragState = StateSizing Then
    
        GetCursorPos pt
   
        DrawDragRect
    
        Select Case m_DragHandle
            Case 0
                m_DragRect.Left = pt.X
                m_DragRect.Top = pt.Y
            Case 1
                m_DragRect.Top = pt.Y
            Case 2
                m_DragRect.Right = pt.X
                m_DragRect.Top = pt.Y
            Case 3
                m_DragRect.Right = pt.X
            Case 4
                m_DragRect.Right = pt.X
                m_DragRect.Bottom = pt.Y
            Case 5
                m_DragRect.Bottom = pt.Y
            Case 6
                m_DragRect.Left = pt.X
                m_DragRect.Bottom = pt.Y
            Case 7
                m_DragRect.Left = pt.X
        End Select

        DrawDragRect
    End If
End Sub

Private Sub doc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
        If m_DragState = StateDragging Or m_DragState = StateSizing Then
           
            DrawDragRect
            'Move control to new location
            m_DragRect.ScreenToTwips m_CurrCtl
            m_DragRect.SetCtrlToRect m_CurrCtl
            'Restore sizing handles
            ShowHandles True
            'Free mouse movement
            ClipCursor ByVal 0&
            'Release mouse capture
            ReleaseCapture
            'Reset drag state
            m_DragState = StateNothing
        End If
    End If
End Sub

Private Sub doc_Resize()
If doc.Height > MDI.Height Then
VScroll1.Max = (doc.Height - MDI.Height)
End If
End Sub



Private Sub Form_Load()
On Error Resume Next
dirty = False
bgclr = vbWhite
vclr = vbRed
tclr = vbBlack
DragInit
runtime = False
Kill App.Path & "\tmp2.tmp"
FileCopy App.Path & "\tmp.tmp", App.Path & "\tmp2.tmp"
Form2.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 Dim i As Integer
If runtime = True Then
ShowHandles False
End If

    If Button = vbLeftButton And runtime = False Then
         dirty = True
          Form7.Visible = False
        If addctrl = "line" Then
        MDI.lneput X, Y
        addctrl = "mouse"
        End If
          If addctrl = "tag" Then
        MDI.hmlput X, Y
        addctrl = "mouse"
        End If
          If addctrl = "image" Then
        MDI.imgput X, Y
        addctrl = "mouse"
        End If
          If addctrl = "text" Then
        MDI.txtput X, Y
        addctrl = "mouse"
        End If
        'Hit test over light-weight (non-windowed) controls
        For i = 0 To (Controls.Count - 1)
             
          
             
     
            If Controls(i).Visible = True Then
                 If Not TypeOf Controls(i) Is Shape Then
               
                  End If
                m_DragRect.SetRectToCtrl Controls(i)
                If m_DragRect.PtInRect(X, Y) Then
                    DragBegin Controls(i)
                    If TypeOf Controls(i) Is Shape Then
                
                    Else
                    
                    End If
                    Exit Sub
                    End If
           
            End If
     
        Next i
        'No control is active
        Set m_CurrCtl = Nothing
        'Hide sizing handles
        ShowHandles False
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nWidth As Single, nHeight As Single
If Button = vbLeftButton Then
    Dim pt As POINTAPI

    If m_DragState = StateDragging Then
         dirty = True
        'Save dimensions before modifying rectangle
        nWidth = m_DragRect.Right - m_DragRect.Left
        nHeight = m_DragRect.Bottom - m_DragRect.Top
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Update drag rectangle coordinates
        m_DragRect.Left = pt.X - m_DragPoint.X
        m_DragRect.Top = pt.Y - m_DragPoint.Y
        m_DragRect.Right = m_DragRect.Left + nWidth
        m_DragRect.Bottom = m_DragRect.Top + nHeight
        'Draw new rectangle
        DrawDragRect
    ElseIf m_DragState = StateSizing Then
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Action depends on handle being dragged
        Select Case m_DragHandle
            Case 0
                m_DragRect.Left = pt.X
                m_DragRect.Top = pt.Y
            Case 1
                m_DragRect.Top = pt.Y
            Case 2
                m_DragRect.Right = pt.X
                m_DragRect.Top = pt.Y
            Case 3
                m_DragRect.Right = pt.X
            Case 4
                m_DragRect.Right = pt.X
                m_DragRect.Bottom = pt.Y
            Case 5
                m_DragRect.Bottom = pt.Y
            Case 6
                m_DragRect.Left = pt.X
                m_DragRect.Bottom = pt.Y
            Case 7
                m_DragRect.Left = pt.X
        End Select
        'Draw new rectangle
        DrawDragRect
    End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
 dirty = True
        If m_DragState = StateDragging Or m_DragState = StateSizing Then
            'Hide drag rectangle
            DrawDragRect
            'Move control to new location
            m_DragRect.ScreenToTwips m_CurrCtl
            m_DragRect.SetCtrlToRect m_CurrCtl
            'Restore sizing handles
            ShowHandles True
            'Free mouse movement
            ClipCursor ByVal 0&
            'Release mouse capture
            ReleaseCapture
            'Reset drag state
            m_DragState = StateNothing
        End If
      
    End If
    
End Sub

Private Sub Form_Paint()
If MDI.Picture11.Tag <> "" Then
Dim X As Single
Dim Y As Single
Y = 0
X = 0
Me.PaintPicture MDI.Picture11.Picture, 0, 0, MDI.Picture11.Width, MDI.Picture11.Height, 0, 0, MDI.Picture11.Width, MDI.Picture11.Height
Do
X = X + MDI.Picture11.Width
Me.PaintPicture MDI.Picture11.Picture, X, Y, MDI.Picture11.Width, MDI.Picture11.Height, 0, 0, MDI.Picture11.Width, MDI.Picture11.Height
If X > Me.Width Then
X = -MDI.Picture11.Width
Y = Y + MDI.Picture11.Height
End If
Loop Until Y > Me.ScaleHeight
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Form1.Left = ((MDI.Width) - Form1.Width - MDI.Toolbar2.Width) / 2
Form1.Top = (MDI.Height - Form1.Height - MDI.Picture13.Height) / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form7
End Sub

Private Sub Image1_Click(Index As Integer)
ccindexx = Index
ctrltype = "image"
End Sub

Private Sub Image1_DblClick(Index As Integer)
On Error Resume Next
Dim str1 As String
Dim str2 As String
indexctrl = Index
Form5.Text1.Text = Image1(indexctrl).ToolTipText
Form5.Text2.Text = Image1(indexctrl).Tag
Form5.Tag = "image"
Form5.Show vbModal
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
Me.PopupMenu MDI.mnuedit
End If
ccindexx = Index
ctrltype = "image"
If runtime = False Then
If Button = vbLeftButton Then
DragBegin Image1(Index)
Form7.Visible = False
End If
 End If
End Sub

Private Sub Label1_Click(Index As Integer)
ccindexx = Index
ctrltype = "text"
End Sub

Private Sub Label1_DblClick(Index As Integer)
On Error Resume Next
Dim str1 As String
Dim str2 As String
indexctrl = Index
Form5.Text2.Text = Label1(indexctrl).Tag
Form5.Tag = "label"
Form5.Show vbModal
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
Me.PopupMenu MDI.mnuedit
End If
ccindexx = Index
ctrltype = "text"
If Button = vbLeftButton And runtime = False Then


Form7.Left = 0
Form7.Top = 0

Form7.Show

MDI.Tag = "label"

changes = True

Form7.Text1.Tag = Index
Form7.Text1.SetFocus
Form7.Text1.Text = Label1(Index).Caption
Form7.Combo2.Text = Label1(Index).FontName
Form7.Combo3.Text = Label1(Index).FontSize
Dim i As Integer
For i = 0 To 2
If i = Label1(Index).Alignment Then Form7.Option1(i).Value = Label1(Index).Alignment
Next i
If Label1(Index).FontBold = True Then
Form7.Check1.Value = 1
Else
Form7.Check1.Value = 0
End If
If Label1(Index).FontItalic = True Then
Form7.Check2.Value = 1
Else
Form7.Check2.Value = 0
End If
If Label1(Index).FontUnderline = True Then
Form7.Check3.Value = 1
Else
Form7.Check3.Value = 0
End If
Form7.Picture4.BackColor = Label1(Index).ForeColor
Form7.Picture5.BackColor = Label1(Index).BackColor

Form7.Check4.Value = Label1(Index).BackStyle
Form7.Text1.SelStart = 0
Form7.Text1.SelLength = Len(Form7.Text1.Text)

Form7.Text1.SetFocus

 End If
 If Button = vbLeftButton Then
DragBegin Label1(Index)
End If
End Sub

Private Sub Label2_Click(Index As Integer)
ccindexx = Index
ctrltype = "line"
End Sub

Private Sub Label2_DblClick(Index As Integer)
On Error GoTo er
MDI.CommonDialog1.ShowColor
Label2(Index).BackColor = MDI.CommonDialog1.Color
Exit Sub
er:
If Err.Number <> 32755 Then
MsgBox Err.Description
End If
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
Me.PopupMenu MDI.mnuedit
End If
ccindexx = Index
ctrltype = "line"
If Button = vbLeftButton Then
DragBegin Label2(Index)
End If
End Sub

Private Sub Label3_DblClick(Index As Integer)
On Error Resume Next
Form6.Show vbModal
ccindexx = Index
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
Me.PopupMenu MDI.mnuedit
End If
ccindexx = Index
ctrltype = "html"
If Button = vbLeftButton Then
DragBegin Label3(Index)
End If
End Sub

Private Sub picHandle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim i As Integer
    Dim rc As RECT

    'Handles should only be visible when a control is selected
    Debug.Assert (Not m_CurrCtl Is Nothing)
    'NOTE: m_DragPoint not used for sizing
    'Save control position in screen coordinates
    m_DragRect.SetRectToCtrl m_CurrCtl
    m_DragRect.TwipsToScreen m_CurrCtl
    'Track index handle
    m_DragHandle = Index
    'Hide sizing handles
    ShowHandles False
    'We need to force handles to hide themselves before drawing drag rectangle
    Refresh
    'Indicate sizing is under way
    m_DragState = StateSizing
    'Show sizing rectangle
    DrawDragRect
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    SetCapture hwnd
    'Limit cursor movement within form
    GetWindowRect hwnd, rc
    ClipCursor rc
End Sub




Private Sub Script_Error()
MsgBox Script.Error.Description & " at " & Script.Error.Text
End Sub
