VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   " About"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2835
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2835
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2850
      Width           =   1260
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   600
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   720
      Picture         =   "frmAbout.frx":0D4A
      Top             =   3540
      Width           =   1500
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2009"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   420
      Width           =   2715
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   2700
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   2700
      Y1              =   2700
      Y2              =   2685
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HMcS Computers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   75
      TabIndex        =   8
      Tag             =   "Application Title"
      Top             =   1020
      Width           =   2655
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.hmcscomputers.com"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   60
      TabIndex        =   7
      Tag             =   "App Description"
      Top             =   2340
      Width           =   2670
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "157 Washington Avenue"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1380
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vandergrift, Pa 15690"
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "724 567-6328"
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   1740
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Howard L. McHenry - Owner"
      Height          =   255
      Left            =   75
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "- Notification Area Behavior -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   180
      Width           =   2715
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "avhlm@comcast.net"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   60
      TabIndex        =   1
      Tag             =   "App Description"
      Top             =   2100
      Width           =   2670
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" ( _
                         ByVal hWnd As Long, _
                         ByVal lpOperation As String, _
                         ByVal lpFile As String, ByVal _
                         lpParameters As String, _
                         ByVal lpDirectory As String, _
                         ByVal nShowCmd As Long) As Long
                
Private Const SW_SHOW = 1

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
                        ByVal hInstance As Long, _
                        ByVal lpCursorName As Long) As Long

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Const IDC_HAND = 32649&
Private Const IDC_ARROW = 32512&

'Enter here
Private Sub Form_Load()
   frmDemo.Enabled = False
   Me.Top = frmDemo.Top + ((frmDemo.Height - Me.Height) / 2)
   Me.Left = frmDemo.Left + (frmDemo.Width - Me.Width) / 2
   
   Label8 = "Version " & App.Major & "." & _
                         App.Minor & "." & _
                         App.Revision
   
End Sub

'ControlBox
Private Sub Form_Unload(Cancel As Integer)
   frmDemo.Enabled = True
   
End Sub

'OK button
Private Sub cmdOK_Click()
   Unload Me

End Sub

'Email
Private Sub Label6_Click()
   Navigator "mailto:avhlm@comcast.net"
   
End Sub

'HMcS website
Private Sub lblDescription_Click()
   Navigator "http://www.hmcscomputers.com"

End Sub

'Navigate
Private Sub Navigator(ByVal NavTo As String)
   Dim hBrowse As Long
   '
   hBrowse = ShellExecute(0&, "open", NavTo, "", "", SW_SHOW)
  
End Sub

'Marble backround for form
Private Sub Form_Paint()
   Dim x As Integer, Y As Integer
   Dim ImgWidth As Integer
   Dim ImgHeight As Integer
   Dim FrmWidth As Integer
   Dim FrmHeight As Integer
   '
   'Use Image1 or Picture1, as appropriate.
   'Use one of the following PaintPicture methods:
   ImgWidth = Image1.Width
   ImgHeight = Image1.Height
   FrmWidth = frmAbout.Width
   FrmHeight = frmAbout.Height
   
   'tile entire form (Method 1)
   For x = 0 To FrmWidth Step ImgWidth
      For Y = 0 To FrmHeight Step ImgHeight
         PaintPicture Image1, x, Y
      Next Y
   Next x
   
   'tile left side (Method 2)
   'For Y = 0 To FrmHeight Step ImgHeight
   '   PaintPicture Image1, 0, Y
   'Next Y
    
End Sub

'Mouse hand cursor
Public Sub SetHandCur(Hand As Boolean)
   If Hand = True Then
      SetCursor LoadCursor(0, IDC_HAND)
   Else
      SetCursor LoadCursor(0, IDC_ARROW)
   End If

End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   SetHandCur True
   
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   SetHandCur True
   
End Sub

Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   SetHandCur True
   
End Sub
Private Sub lblDescription_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   SetHandCur True
   
End Sub




