VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tray Balloon Tip Test"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7035
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Timer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   10
      Top             =   900
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3000
      Top             =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "This Frame Hides at Form_Load"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   150
      TabIndex        =   5
      Top             =   1425
      Width           =   4665
      Begin VB.PictureBox Picture1 
         Height          =   315
         Left            =   975
         Picture         =   "Form1.frx":2CFA
         ScaleHeight     =   255
         ScaleWidth      =   330
         TabIndex        =   8
         Top             =   1350
         Width           =   390
      End
      Begin VB.PictureBox Picture2 
         Height          =   315
         Left            =   1575
         Picture         =   "Form1.frx":59F4
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   1350
         Width           =   315
      End
      Begin VB.PictureBox Picture3 
         Height          =   690
         Left            =   1950
         Picture         =   "Form1.frx":86EE
         ScaleHeight     =   630
         ScaleWidth      =   630
         TabIndex        =   6
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"Form1.frx":8FB8
         Height          =   615
         Left            =   75
         TabIndex        =   9
         Top             =   450
         Width           =   4440
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Test No Icon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5400
      TabIndex        =   4
      Top             =   1725
      Width           =   1440
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Test Error"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5400
      TabIndex        =   3
      Top             =   1200
      Width           =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test Warning"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5400
      TabIndex        =   2
      Top             =   675
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5400
      TabIndex        =   1
      Top             =   150
      Width           =   1440
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":904A
      Top             =   150
      Width           =   4965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer ' for the timer.. you can ditch the timer, it's all about fluff anyway
 
Private m_IconData As NOTIFYICONDATA
 
Private Sub Command1_Click()  'not currently used

'Set up the structure
' IF szInfoTitle and szInfo are not terminated with 'null' the box is huge!
For i = 1 To 10
txt = txt & "This is a sample piece of text"
Next i
    With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Picture1.Picture 'Me.Icon
        .szTip = "Your Message Here" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "This is the result" & vbCrLf & "Pretty cool huh?" & Chr(0)
        '.szInfo = txt
        .szInfoTitle = "Button Clicked" & Chr(0)
        .dwInfoFlags = 1 'NIIF_ERROR
        .uTimeout = 3000
   End With
  
        Shell_NotifyIcon NIM_MODIFY, m_IconData

End Sub

Private Sub Command2_Click()
'Set up the structure
' IF szInfoTitle and szInfo are not terminated with 'null' the box is huge!
For i = 1 To 10
txt = txt & "This is a sample piece of text"
Next i
    With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Picture2.Picture 'Me.Icon
        .szTip = "Your Message Here" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "The Info button was clicked, notice the 'Info' Icon" & Chr(0)
      
        .szInfoTitle = "Info Button Clicked" & Chr(0)
        .dwInfoFlags = NIIF_INFO
        .uTimeout = 3000
   End With
  
        Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Private Sub Command3_Click()

' IF szInfoTitle and szInfo are not terminated with 'null' the box is huge!
For i = 1 To 10
txt = txt & "This is a sample piece of text"
Next i
    With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Picture1.Picture 'Me.Icon
        .szTip = "Your Message Here" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = txt '"This is the result" & vbCrLf & "Pretty cool huh?" & Chr(0)
        '.szInfo = txt
        .szInfoTitle = "Warning Button Clicked" & Chr(0)
        .dwInfoFlags = NIIF_WARNING
        .uTimeout = 3000
   End With
  
        Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Private Sub Command4_Click()
'Set up the structure
' IF szInfoTitle and szInfo are not terminated with 'null' the box is huge!
For i = 1 To 10
txt = txt & "This is a sample piece of text"
Next i
    With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Picture3.Picture 'Me.Icon
        .szTip = "Your Message Here" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "This is the result" & vbCrLf & "Pretty cool huh?" & Chr(0)
        '.szInfo = txt
        .szInfoTitle = "Error Button Clicked" & Chr(0)
        .dwInfoFlags = NIIF_ERROR
        .uTimeout = 3000
   End With
  
        Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Private Sub Command5_Click()
'Set up the structure
' IF szInfoTitle and szInfo are not terminated with 'null' the box is huge!
For i = 1 To 10
txt = txt & "This is a sample piece of text"
Next i
    With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Picture1.Picture
        .szTip = "Your Message Here" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "This is the result" & vbCrLf & "Pretty cool huh?" & Chr(0)
        '.szInfo = txt 'uncomment this to fill the box with junk
        .szInfoTitle = "No Icon Button Clicked" & Chr(0)
        .dwInfoFlags = NIIF_NONE
        .uTimeout = 3000
   End With
  
        Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Private Sub Form_Load()
Frame1.Visible = False
    With m_IconData
        .cbSize = Len(m_IconData)
        .hWnd = Me.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Your Message Here" & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "This is a balloon-style tool-tip" & Chr(0)
        .szInfoTitle = "This is a balloon-tip title" & Chr(0)
        .dwInfoFlags = NIIF_WARNING
         End With
   Shell_NotifyIcon NIM_ADD, m_IconData
End Sub


Private Sub Form_Unload(Cancel As Integer)

'This deletes the icon in the tray for program termination
Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

Private Sub Timer1_Timer()
If Not Check1.Value = Checked Then Exit Sub
counter = counter + 1
If counter > 4 Then counter = 1
Select Case counter
Case 1
Command2_Click
Case 2
Command3_Click
Case 3
Command4_Click
Case 4
Command5_Click
Case Else
End Select

End Sub
