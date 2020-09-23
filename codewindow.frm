VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Code Window"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6165
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1320
      Width           =   6135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vb code"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hyperlink:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Me.Tag = "label" Then
tcde(Text1.Tag) = Text2.Text
Unload Me
End If
If Me.Tag = "image" Then
icde(Text1.Tag) = Text2.Text
Unload Me
End If
End Sub

