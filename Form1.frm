VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNote 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Note"
   ClientHeight    =   3030
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   4140
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":000C
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEvent 
      Interval        =   1000
      Left            =   0
      Top             =   2640
   End
   Begin MSComCtl2.DTPicker myTime 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   65535
      CalendarForeColor=   255
      CalendarTitleBackColor=   65535
      CalendarTitleForeColor=   255
      CalendarTrailingForeColor=   65535
      Format          =   22806530
      CurrentDate     =   37216
   End
   Begin MSComCtl2.DTPicker myDate 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   65535
      CalendarForeColor=   0
      CalendarTitleBackColor=   65535
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   65535
      Format          =   22806529
      CurrentDate     =   37216
   End
   Begin VB.TextBox txtNote 
      BackColor       =   &H0000FFFF&
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
      Height          =   1935
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblMinimise 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   200
      X2              =   208
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   135
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   216
      X2              =   224
      Y1              =   24
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   216
      X2              =   224
      Y1              =   16
      Y2              =   24
   End
   Begin VB.Menu mnopt 
      Caption         =   "options"
      Begin VB.Menu mnuSetDateTime 
         Caption         =   "Set Date/Time To Show Note"
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStayOnTop 
         Caption         =   "Stay On Top"
      End
      Begin VB.Menu mnuHideMin 
         Caption         =   "Hide Note When Minimised"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim TimeChanged As Boolean

Private Sub mnuHideMin_Click()
    
    If mnuHideMin.Checked = False Then
        mnuHideMin.Checked = True
    Else
        mnuHideMin.Checked = False
    End If
    
End Sub

Private Sub mnuSetDateTime_Click()
    If mnuSetDateTime.Checked = False Then
        'make date time setters visible
        myTime.Visible = True
        myDate.Visible = True
        mnuSetDateTime.Checked = True
        'set textbox height
        txtNote.Height = 105
    Else
        myTime.Visible = False
        myDate.Visible = False
        mnuSetDateTime.Checked = False
        'turn off timer if it's on
        tmrEvent.Enabled = False
        'set textbox height
        txtNote.Height = 129
    End If
End Sub

Private Sub mnuStayOnTop_Click()
    
    If mnuStayOnTop.Checked = False Then
        rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
        mnuStayOnTop.Checked = True
    Else
        rtn = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
        mnuStayOnTop.Checked = False
    End If
    
    txtNote.SetFocus
    
End Sub

Private Sub Form_Load()

myDate.Value = Date
tmrEvent.Enabled = False


If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)
End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Label1_Click()
    Me.PopupMenu Me.mnopt
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblMinimise_Click()
    
    If TimeChanged = True Then
        TimeChanged = False
        tmrEvent.Enabled = True
    End If
    
    If mnuHideMin.Checked = False Then
        Me.WindowState = vbMinimized
    Else
        tmrEvent.Enabled = True
        Me.Visible = False
    End If
    
End Sub

Private Sub myTime_Change()
    TimeChanged = True
End Sub

Private Sub tmrEvent_Timer()
    'MsgBox Format(myTime, "hh:mm:ss")
    If myDate = Date And Format(myTime, "hh:mm") = Format(Time, "hh:mm") Then
        Me.Show
        rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
        tmrEvent.Enabled = False
    End If
End Sub
