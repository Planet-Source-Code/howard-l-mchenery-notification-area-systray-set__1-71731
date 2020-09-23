VERSION 5.00
Begin VB.Form frmNotifyArea 
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11985
   Icon            =   "frmNotifyArea.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmNotifyArea.frx":0D4A
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu Dummy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmNotifyArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Start        Length   Type   Type_Data
'--------------------------------------------------
'0                20   Data
'&H0014 / 20     522   Path   Unicode
'&H021C / 540     16   Data
'&H022C / 556    526   Title  Unicode
'--------------------------------------------------
'Record Length  1084

'Set manually, have to logoff / logon
'Set in registry, stop and restart explorer works

Private aByte() As Byte

'Enter here
Private Sub Form_Load()
   Dim lRet    As Boolean
   Dim x       As Long
   Dim cTxt    As String
   Dim cBehave As String
   '
   frmDemo.Enabled = False
   
   Me.Caption = " Notification Area - " & _
                App.Major & "." & _
                App.Minor & "." & _
                App.Revision
   
   lRet = REG_GetBinary_BYTE(HKEY_CURRENT_USER, _
                             "Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify", _
                             "IconStreams", _
                             aByte())
   
   Text1.Text = vbCrLf & "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\TrayNotify\IconStreams (REG_BINARY)" & vbCrLf & vbCrLf
   Text1.Text = Text1.Text & "Start           Length   Values          Data_Type" & vbCrLf
   Text1.Text = Text1.Text & "-------------------------------------------------------" & vbCrLf
   Text1.Text = Text1.Text & "&H0000 /   0        20   Data Area       Byte" & vbCrLf
   Text1.Text = Text1.Text & "&H0014 /  20       522   FileSpec        Unicode String" & vbCrLf
   Text1.Text = Text1.Text & "&H021C / 540        16   Data Area       Byte" & vbCrLf
   Text1.Text = Text1.Text & "&H022C / 556       526   Title (ToolTip) Unicode String" & vbCrLf
   Text1.Text = Text1.Text & "-------------------------------------------------------" & vbCrLf
   Text1.Text = Text1.Text & "Record Length     1084" & vbCrLf & vbCrLf
   Text1.Text = Text1.Text & "&H0220: Hide when Inactive = 0, Always Hide = 1, Always Show = 0" & vbCrLf
   Text1.Text = Text1.Text & "&H0224: Hide when Inactive = 0, Always Hide = 1, Always Show = 2" & vbCrLf & vbCrLf
   Text1.Text = Text1.Text & Chr(34) & "Past Item" & Chr(34) & " if process is not running." & vbCrLf
   Text1.Text = Text1.Text & "Change registry, stop and restart " & Chr(34) & "Explorer" & Chr(34) & " to set changes." & vbCrLf
   Text1.Text = Text1.Text & "Remove " & Chr(34) & "Past Items," & Chr(34) & " delete " & Chr(34) & "IconStreams" & Chr(34) & " and " & Chr(34) & "PastIconsStream" & Chr(34) & " registry values and logoff." & vbCrLf
   Text1.Text = Text1.Text & "Right click " & Chr(34) & "Start" & Chr(34) & ", " & Chr(34) & "Properties" & Chr(34) & ", " & Chr(34) & "Taskbar" & Chr(34) & ", " & Chr(34) & "Customize" & Chr(34) & " to manually view settings." & vbCrLf & vbCrLf
   
   For x = 0 To UBound(aByte) Step 1084
      'ToolTip &H022C to &H043C
      cTxt = GetText(x, 556, 1084)
      If cTxt <> "" Then
         Text1.Text = Text1.Text & Right("  " & Str(Int(x / 1083 + 1)), 2) & ". " & cTxt & vbCrLf
      
         'Path &H0014 to &H020A
         cTxt = GetText(x, 20, 522)
         Text1.Text = Text1.Text & "    " & cTxt & vbCrLf
         
         'Data &H0000 to &H0014
         cTxt = GetData(x, 0, 19)
         Text1.Text = Text1.Text & "    &H0000 " & cTxt & vbCrLf
         
         'Data &H021C to &H022C
         cTxt = GetData(x, 540, 555)
         Text1.Text = Text1.Text & "    &H021C " & cTxt & vbCrLf
         
         'Data &H0224 HI = 0, AH = 1, AS = 2
         cTxt = GetData(x, &H220, 548)

         Select Case Val(cTxt)
            Case 0
               cBehave = "Hide when Inactive"
               
            Case 100000001
               cBehave = "Always Hide"
               
            Case 2
               cBehave = "Always Show"
               
         End Select
         
         'Test for Past Item
         If Not IsProcessRun(GetText(x, 20, 522)) Then
            cBehave = "Past Item - " & cBehave
         End If
         
         Text1.Text = Text1.Text & "    &H0220 " & cTxt & " &H0224 - " & cBehave & vbCrLf
         
         Text1.Text = Text1.Text & vbCrLf
      End If
   Next
   
End Sub

'Resize
Private Sub Form_Resize()
   Text1.Width = Me.ScaleWidth
   Text1.Height = Me.ScaleHeight
   
End Sub

'ControlBox
Private Sub Form_Unload(Cancel As Integer)
   frmDemo.Enabled = True
   
End Sub

'Get unicode text from byte array
Private Function GetText(ByVal nRec As Long, ByVal nStrt As Long, ByVal nStop As Long) As String
   Dim cTxt As String
   Dim i    As Long
   Dim Y    As Long
   '
   For Y = nStrt To nStop
      i = nRec + Y
      If i > UBound(aByte) Then
         Exit For
      End If
      cTxt = cTxt & Chr(aByte(i))
   Next
   cTxt = StrConv(cTxt, vbFromUnicode)
   cTxt = TrimWithoutPrejudice(cTxt)
   
   GetText = cTxt
   
End Function

'Get hex from byte array
Private Function GetData(ByVal nRec As Long, ByVal nStrt As Long, ByVal nStop As Long) As String
   Dim cTxt As String
   Dim i    As Long
   Dim Y    As Long
   '
   For Y = nStrt To nStop
      i = nRec + Y
      If i > UBound(aByte) Then
         Exit For
      End If
      If cTxt = "" Then
         cTxt = cTxt & Right("00" & Hex(aByte(i)), 2)
      Else
         cTxt = cTxt & " " & Right("00" & Hex(aByte(i)), 2)
      End If
   Next
   
   GetData = cTxt
   
End Function

'Eliminate non-printable characters
Private Function TrimWithoutPrejudice(ByVal InputString As String) As String
   Dim sAns  As String
   Dim sWkg  As String
   Dim sChar As String
   Dim lLen  As Long
   Dim lCtr  As Long
   '
   sAns = InputString
   lLen = Len(InputString)
   
   If lLen > 0 Then
      'Ltrim
      For lCtr = 1 To lLen
         sChar = Mid(sAns, lCtr, 1)
         If Asc(sChar) > 32 Then
            Exit For
         End If
      Next
   
      sAns = Mid(sAns, lCtr)
      lLen = Len(sAns)
   
      'Rtrim
      If lLen > 0 Then
         For lCtr = lLen To 1 Step -1
            sChar = Mid(sAns, lCtr, 1)
            If Asc(sChar) > 32 Then
               Exit For
            End If
         Next
      End If
      sAns = Left$(sAns, lCtr)
   End If
   
   TrimWithoutPrejudice = sAns

End Function

'Is process running
Private Function IsProcessRun(ByVal cFileSpec As String) As Boolean
   Dim Process As Object
   '
   cFileSpec = Right(cFileSpec, Len(cFileSpec) - InStrRev(cFileSpec, "\"))
   For Each Process In GetObject("winmgmts:"). _
      ExecQuery("select * from Win32_Process where name='" & cFileSpec & "'")
      IsProcessRun = True
   Next
      
End Function

'File > Save menu
Private Sub mnuSave_Click()
   Dim nFno As Integer
   '
   nFno = FreeFile
   Open App.Path & "\Notification Area.txt" For Output As #nFno
   Print #nFno, Text1.Text
   Close #nFno
   MsgBox "Text Saved to File in Current Folder", vbInformation
   
End Sub

'File > Exit menu
Private Sub mnuExit_Click()
   Unload Me
   
End Sub































