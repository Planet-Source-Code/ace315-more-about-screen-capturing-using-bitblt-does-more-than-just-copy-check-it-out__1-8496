VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Screen Captureing"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   3735
      TabIndex        =   8
      Top             =   2160
      Width           =   3735
      Begin VB.OptionButton Option7 
         Caption         =   "Flip Screen Vertical"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Flip Screen Horizontal"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Flip Screen"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   3735
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Brighten Screen"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Darken"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Bad colors"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Invert the screen"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Keep screen normal"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Capture"
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Screen Options"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   3720
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   0
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Line Line4 
      X1              =   3720
      X2              =   3720
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Line Line3 
      X1              =   3720
      X2              =   2040
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2040
      Y1              =   480
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   2040
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fliphorizontal As Boolean, flipvertical As Boolean, thechange  'declare the variables
Private Sub Check1_Click()
Option6.Enabled = Not Option6.Enabled 'disable/enable flip screen options
Option7.Enabled = Option6.Enabled
If Option6.Enabled = False Then
fliphorizontal = False 'set variables to false
flipvertical = False
End If
End Sub
Private Sub Command1_Click()
Form2.Picture1.Cls 'Clear picture
DumpToWindow Form2.Picture1, thechange, fliphorizontal, flipvertical
Form2.Show 'show the form
End Sub
Private Sub Form_Load()
fliphorizontal = False 'set variable to correct value
flipvertical = False
thechange = SRCCOPY
With Form2 'set the size of the form and picture in it
.Top = 0
.Left = 0
.Width = Screen.Width
.Height = Screen.Height
.Picture1.Height = Screen.Height
.Picture1.Width = Screen.Width
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Option1_Click()
thechange = SRCCOPY 'change variable
End Sub
Private Sub Option2_Click()
thechange = SRCINVERT 'change variable
End Sub
Private Sub Option3_Click()
thechange = SRCAND 'change variable
End Sub
Private Sub Option4_Click()
thechange = SRCERASE 'change variable
End Sub
Private Sub Option5_Click()
thechange = SRCPAINT 'change variable
End Sub
Private Sub Option6_Click()
fliphorizontal = True 'change variables
flipvertical = False
End Sub
Private Sub Option7_Click()
fliphorizontal = False 'change variables
flipvertical = True
End Sub
