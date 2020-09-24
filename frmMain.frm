VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sticky Notes"
   ClientHeight    =   45
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   1590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   45
   ScaleWidth      =   1590
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuNew 
         Caption         =   "New Sticky Note"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show All Sticky Notes"
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    
    Dim strMyToolTip As String
    
    strMyToolTip = "Sticky Notes"
    Call AddSystray(Me, strMyToolTip)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errorH
  Dim rtn As Long

  'this procedure receives the callbacks from the
  'System Tray icon and pops up the menu if the right
  'button is clicked.  I left the other button options
  'there, incase you want to have other options...

    'just incase we are trying to process something, let's block this out...
    'the value of X will vary depending on the scalemode setting
    If Me.ScaleMode = vbPixels Then
        rtn = X
      Else
        rtn = X / Screen.TwipsPerPixelX
    End If

    Select Case rtn
      Case WM_LBUTTONDOWN         '= &H201 - Left Button down
        'nothing happens, yet
      Case WM_LBUTTONUP           '= &H202 - Left Button up
        'nothing happens, yet
      Case WM_LBUTTONDBLCLK       '= &H203 - Left Double-click
        'nothing happens, yet
      Case WM_RBUTTONDOWN         '= &H204 - Right Button down
        'nothing happens, yet
      Case WM_RBUTTONUP           '= &H205 - Right Button up
        SetForegroundWindow Me.hWnd
        Me.PopupMenu Me.mnuOptions
      Case WM_RBUTTONDBLCLK       '= &H206 - Right Double-click
        'nothing happens, yet
    End Select

Exit Sub

errorH:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RemoveSystray
    End
End Sub

Private Sub mnuNew_Click()

   Dim TotalForms As Integer
    ' We Fill TotalForms with the Total Number of forms +1
    TotalForms = Forms.Count + 1
    ' we call our subroutine to create the from.
    ' Totalforms is converted to a string and used for in name and caption
    Dform.CreateForm "Form" + CStr(TotalForms), "Sticky Note"

End Sub

Private Sub mnuExit_Click()

    Unload Me

End Sub

Private Sub mnuShowAll_Click()
    For i = 1 To Forms.Count - 1
        Forms(i).Show
    Next i
End Sub
