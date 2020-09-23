VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   6465
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Width           =   6315
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Behavior to: Hide when Inactive"
      Height          =   495
      Left            =   1620
      TabIndex        =   2
      Top             =   1800
      Width           =   3195
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Behavior to: Always Hide"
      Height          =   495
      Left            =   1620
      TabIndex        =   1
      Top             =   1200
      Width           =   3195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Behavior to: Always Show"
      Height          =   495
      Left            =   1620
      TabIndex        =   0
      Top             =   600
      Width           =   3195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   2340
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   " File Spec:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuDiscovData 
         Caption         =   "Discovety Data"
      End
      Begin VB.Menu Dummy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Enter here
Private Sub Form_Load()
   Me.Caption = " Notification Area (Systray) Behavior Set - " & _
                App.Major & "." & _
                App.Minor & "." & _
                App.Revision
   
   Text1.Text = "C:\Program Files\Call Trace\ctrace.exe"
   
   Label2.Caption = BehaviorGet(Text1.Text)
   
End Sub

'Set Behavior to: Always Show button
Private Sub Command1_Click()
   If BehaviorSet(Text1.Text) Then
      MsgBox "Notification Area Behavior Successfully Set", vbInformation
   Else
      MsgBox "Problem Setting Notification Area Behavior", vbCritical
   End If
   
   RefreshLabel
   
End Sub

'Set Behavior to: Always Hide button
Private Sub Command2_Click()
   If BehaviorSet(Text1.Text, BHV_ALWHIDES) Then
      MsgBox "Notification Area Behavior Successfully Set", vbInformation
   Else
      MsgBox "Problem Setting Notification Area Behavior", vbCritical
   End If
   
   RefreshLabel
   
End Sub

'Set Behavior to: Hide when Inactive button
Private Sub Command3_Click()
   If BehaviorSet(Text1.Text, BHV_HIDINACT) Then
      MsgBox "Notification Area Behavior Successfully Set", vbInformation
   Else
      MsgBox "Problem Setting Notification Area Behavior", vbCritical
   End If
   
   RefreshLabel
   
End Sub

'Change text in file spec textbox
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
   RefreshLabel
   
End Sub

'Refresh label
Private Sub RefreshLabel()
   Label2.Caption = BehaviorGet(Text1.Text)
   
End Sub

'File > About menu
Private Sub mnuAbout_Click()
   frmAbout.Show
   
End Sub

'File > Discovery Data menu
Private Sub mnuDiscovData_Click()
   frmNotifyArea.Show
   
End Sub

'File > Exit menu
Private Sub mnuExit_Click()
   Unload Me
   
End Sub



















