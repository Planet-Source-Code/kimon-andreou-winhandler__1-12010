VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Handler"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShowChildWindows 
      Caption         =   "&Show Children"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6060
      TabIndex        =   62
      Top             =   8340
      Width           =   1425
   End
   Begin VB.Frame fr 
      Caption         =   "SetWindowPos Parameters"
      Enabled         =   0   'False
      Height          =   2805
      Index           =   4
      Left            =   0
      TabIndex        =   16
      Top             =   5370
      Width           =   9135
      Begin VB.OptionButton optWindow 
         Caption         =   "Child"
         Height          =   195
         Index           =   1
         Left            =   8040
         TabIndex        =   50
         ToolTipText     =   "Execute function on child window."
         Top             =   660
         Width           =   915
      End
      Begin VB.OptionButton optWindow 
         Caption         =   "Owner"
         Height          =   195
         Index           =   0
         Left            =   8040
         TabIndex        =   49
         ToolTipText     =   "Execute function on window selected above."
         Top             =   450
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   405
         Left            =   8040
         TabIndex        =   52
         Top             =   2190
         Width           =   1005
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         Height          =   405
         Left            =   8040
         TabIndex        =   51
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Frame fr 
         Caption         =   "hWndInsertAfter"
         Height          =   1575
         Index           =   7
         Left            =   60
         TabIndex        =   27
         Top             =   1170
         Width           =   2895
         Begin VB.TextBox txthWndInsertAfter 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1140
            TabIndex        =   33
            Top             =   1260
            Width           =   705
         End
         Begin VB.OptionButton optHWndInsertAfter 
            Caption         =   "hWnd"
            Height          =   225
            Index           =   4
            Left            =   90
            TabIndex        =   32
            ToolTipText     =   "Puts window after specified window."
            Top             =   1290
            Width           =   885
         End
         Begin VB.OptionButton optHWndInsertAfter 
            Caption         =   "HWND_TOPMOST"
            Height          =   225
            Index           =   3
            Left            =   90
            TabIndex        =   31
            ToolTipText     =   "Places the window above all non-topmost windows. "
            Top             =   1050
            Width           =   2055
         End
         Begin VB.OptionButton optHWndInsertAfter 
            Caption         =   "HWND_TOP"
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   30
            ToolTipText     =   "Places the window at the top of the Z order."
            Top             =   810
            Width           =   2055
         End
         Begin VB.OptionButton optHWndInsertAfter 
            Caption         =   "HWND_NOTOPMOST"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   29
            ToolTipText     =   "Places the window above all non-topmost windows (that is, behind all topmost windows). "
            Top             =   570
            Width           =   2055
         End
         Begin VB.OptionButton optHWndInsertAfter 
            Caption         =   "HWND_BOTTOM"
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   28
            ToolTipText     =   "Places the window at the bottom of the Z order. "
            Top             =   330
            Width           =   1635
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Window Size"
         Height          =   945
         Index           =   0
         Left            =   1500
         TabIndex        =   22
         Top             =   210
         Width           =   1455
         Begin VB.TextBox txtWindowWidth 
            Height          =   285
            Left            =   690
            MaxLength       =   4
            TabIndex        =   24
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtWindowHeight 
            Height          =   285
            Left            =   690
            MaxLength       =   4
            TabIndex        =   26
            Text            =   "0"
            Top             =   570
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Width"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   23
            Top             =   270
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Height"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   465
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Screen Position"
         Height          =   945
         Index           =   6
         Left            =   30
         TabIndex        =   17
         Top             =   210
         Width           =   1455
         Begin VB.TextBox txtScreenTop 
            Height          =   285
            Left            =   690
            MaxLength       =   4
            TabIndex        =   21
            Text            =   "0"
            Top             =   570
            Width           =   495
         End
         Begin VB.TextBox txtScreenLeft 
            Height          =   285
            Left            =   690
            MaxLength       =   4
            TabIndex        =   19
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Top"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   285
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Left"
            Height          =   195
            Index           =   5
            Left            =   150
            TabIndex        =   18
            Top             =   270
            Width           =   270
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Flags"
         Height          =   2535
         Index           =   5
         Left            =   2970
         TabIndex        =   34
         Top             =   210
         Width           =   4995
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_DEFERERASE"
            Height          =   285
            Index           =   13
            Left            =   2550
            TabIndex        =   48
            ToolTipText     =   "Prevents generation of the WM_SYNCPAINT message. "
            Top             =   1200
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOSENDCHANGING"
            Height          =   285
            Index           =   12
            Left            =   2550
            TabIndex        =   47
            ToolTipText     =   "Prevents the window from receiving the WM_WINDOWPOSCHANGING message."
            Top             =   960
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_SHOWWINDOW"
            Height          =   285
            Index           =   11
            Left            =   2550
            TabIndex        =   46
            ToolTipText     =   "Displays the window."
            Top             =   720
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOZORDER"
            Height          =   285
            Index           =   10
            Left            =   2550
            TabIndex        =   45
            ToolTipText     =   "Retains the current Z order (ignores the hWndInsertAfter parameter)."
            Top             =   480
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOSIZE"
            Height          =   285
            Index           =   9
            Left            =   2550
            TabIndex        =   44
            ToolTipText     =   "Retains the current size (ignores the cx and cy parameters)."
            Top             =   240
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOREPOSITION"
            Height          =   285
            Index           =   8
            Left            =   180
            TabIndex        =   43
            ToolTipText     =   "Same as the SWP_NOOWNERZORDER flag."
            Top             =   2160
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOREDRAW"
            Height          =   285
            Index           =   7
            Left            =   180
            TabIndex        =   42
            ToolTipText     =   "Does not redraw changes. If this flag is set, no repainting of any kind occurs. "
            Top             =   1920
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOOWNERZORDER"
            Height          =   285
            Index           =   6
            Left            =   180
            TabIndex        =   41
            ToolTipText     =   "Does not change the owner window's position in the Z order."
            Top             =   1680
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOMOVE"
            Height          =   285
            Index           =   5
            Left            =   180
            TabIndex        =   40
            ToolTipText     =   "Retains the current position (ignores the X and Y parameters)."
            Top             =   1440
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOCOPYBITS"
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   39
            ToolTipText     =   "Discards the entire contents of the client area. "
            Top             =   1200
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_NOACTIVATE"
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   38
            ToolTipText     =   "Does not activate the window. "
            Top             =   960
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_HIDEWINDOW"
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   37
            ToolTipText     =   "Hides the window."
            Top             =   720
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_DRAWFRAME"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   36
            ToolTipText     =   "Draws a frame (defined in the window's class description) around the window."
            Top             =   480
            Width           =   2325
         End
         Begin VB.CheckBox chkSetWindowPosFlags 
            Caption         =   "SWP_FRAMECHANGED"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   35
            ToolTipText     =   "Applies new frame styles set using the SetWindowLong function. "
            Top             =   240
            Width           =   2325
         End
      End
   End
   Begin VB.PictureBox status 
      Align           =   2  'Align Bottom
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   9990
      TabIndex        =   58
      Top             =   8880
      Width           =   10050
      Begin VB.Label lblOwnerHwnd 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4785
         TabIndex        =   61
         Top             =   30
         Width           =   60
      End
      Begin VB.Label lblCurrenthWnd 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   9885
         TabIndex        =   60
         Top             =   30
         Width           =   45
      End
      Begin VB.Label lblWindowsFound 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   30
         TabIndex        =   59
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdEnumWinProc 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7740
      TabIndex        =   57
      Top             =   8340
      Width           =   1425
   End
   Begin VB.ListBox lstEnumWindows 
      Height          =   4545
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   10035
   End
   Begin VB.Frame fr 
      Enabled         =   0   'False
      Height          =   735
      Index           =   8
      Left            =   0
      TabIndex        =   53
      Top             =   8130
      Width           =   5655
      Begin VB.CommandButton cmdSetNewTitle 
         Caption         =   "Set Text"
         Height          =   405
         Left            =   4620
         TabIndex        =   56
         Top             =   210
         Width           =   795
      End
      Begin VB.TextBox txtNewWindowTitle 
         Height          =   285
         Left            =   1170
         TabIndex        =   55
         Top             =   270
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Window Text"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   54
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.Frame fr 
      Height          =   645
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   4710
      Width           =   4005
      Begin VB.CommandButton cmdSendMessage 
         Caption         =   "Send Message"
         Enabled         =   0   'False
         Height          =   405
         Left            =   2520
         TabIndex        =   6
         Top             =   150
         Width           =   1335
      End
      Begin VB.ComboBox cboMessages 
         Height          =   315
         Left            =   60
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   2385
      End
   End
   Begin VB.Frame fr 
      Height          =   645
      Index           =   2
      Left            =   4020
      TabIndex        =   7
      Top             =   4710
      Width           =   2115
      Begin VB.CommandButton cmdSetTopMost 
         Caption         =   "Set Top"
         Enabled         =   0   'False
         Height          =   405
         Left            =   60
         TabIndex        =   8
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton cmdUnsetTopMost 
         Caption         =   "Unset Top"
         Enabled         =   0   'False
         Height          =   405
         Left            =   1080
         TabIndex        =   9
         Top             =   150
         Width           =   945
      End
   End
   Begin VB.Frame fr 
      Height          =   645
      Index           =   3
      Left            =   6150
      TabIndex        =   10
      Top             =   4710
      Width           =   2685
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Text            =   "0"
         Top             =   300
         Width           =   555
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   900
         TabIndex        =   14
         Text            =   "0"
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton cmdSetSize 
         Caption         =   "Set Size"
         Height          =   405
         Left            =   1650
         TabIndex        =   15
         Top             =   180
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   11
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Height"
         Height          =   195
         Index           =   4
         Left            =   930
         TabIndex        =   13
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Window Title"
      Height          =   195
      Index           =   2
      Left            =   3930
      TabIndex        =   2
      Top             =   30
      Width           =   930
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "hWnd"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   30
      Width           =   435
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Class"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Local variables
Private SetWindowPosFlags As Long
Private hWndInsertAfter As Long
Private dwFLAGS(13) As Long
Private IamActivated As Boolean

'Executes SetWindowProc()
Private Sub cmdExecute_Click()
Dim flag As Long
Dim cnt As Integer
Dim handle As Long

'Checks which flags are set
For cnt = 0 To 13
    If chkSetWindowPosFlags(cnt).Value = 1 Then
        flag = flag Or dwFLAGS(cnt)
    End If
Next cnt

'Check the value for hWndInsertAfter
If hWndInsertAfter = -666 Then
    hWndInsertAfter = Val(txthWndInsertAfter.Text)
    If hWndInsertAfter = 0 Then
        MsgBox "You have to select something for hWndInsertAfter!", vbExclamation, "Missing argument!"
        hWndInsertAfter = -666
    End If
End If

'See if the function will be called against an owner window or a child window
handle = IIf(optWindow(0).Value = True, CurrenthWnd, Childhwnd)

'Call the function
Call SetWindowPos(handle, hWndInsertAfter, Val(txtScreenLeft.Text), _
    Val(txtScreenTop.Text), Val(txtWindowWidth.Text), Val(txtWindowHeight.Text), flag)
    
End Sub

'Reset all the settings for SetWindowProc
Private Sub cmdReset_Click()
Dim cnt As Integer

For cnt = 0 To 13
    chkSetWindowPosFlags(cnt).Value = 0
Next cnt

txtScreenTop.Text = 0
txtScreenLeft.Text = 0
txtWindowWidth.Text = 0
txtWindowHeight.Text = 0
End Sub

'Send the selected message
Private Sub cmdSendMessage_Click()
Call SendMessage(CurrenthWnd, GetMessageValue(cboMessages.Text), 0&, ByVal 0&)
End Sub

'Set a new title for the selected window
Private Sub cmdSetNewTitle_Click()
    Call SetWindowText(CurrenthWnd, txtNewWindowTitle.Text)
End Sub


Private Sub cmdSetSize_Click()
Dim wdth As Long
Dim hght As Long

wdth = Val(txtWidth.Text)
hght = Val(txtHeight.Text)
If (wdth = 0) Or (hght = 0) Then
    MsgBox "You must specify a height AND a width!", vbExclamation, "Error!"
    Exit Sub
End If

Call SetWindowPos(CurrenthWnd, HWND_TOP, 0, 0, wdth, hght, SWP_NOMOVE Or SWP_NOACTIVATE Or SWP_NOSENDCHANGING)
End Sub

Private Sub cmdSetTopMost_Click()
Call SetWindowPos(CurrenthWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub


Private Sub cmdShowChildWindows_Click()
    frmChildWindows.Show
    frmChildWindows.RefreshMe
optWindow(1).Enabled = True
End Sub

Private Sub cmdUnsetTopMost_Click()
Call SetWindowPos(CurrenthWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub Form_Activate()
If Not IamActivated Then IamActivated = True
End Sub

Private Sub Form_Load()
ReDim TabArray(0 To 2) As Long
   
TabArray(0) = 0
TabArray(1) = -162
TabArray(2) = 174
   
'clear any existing tabs
Call SendMessage(lstEnumWindows.hWnd, LB_SETTABSTOPS, 0&, ByVal 0&)
   
'set list tabstops
Call SendMessage(lstEnumWindows.hWnd, LB_SETTABSTOPS, 3&, TabArray(0))


InitializeCombo
cmdEnumWinProc_Click
   
'Initialize the array where the SetWindowProc() flags are stored.
dwFLAGS(0) = SWP_FRAMECHANGED
dwFLAGS(1) = SWP_DRAWFRAME
dwFLAGS(2) = SWP_HIDEWINDOW
dwFLAGS(3) = SWP_NOACTIVATE
dwFLAGS(4) = SWP_NOCOPYBITS
dwFLAGS(5) = SWP_NOMOVE
dwFLAGS(6) = SWP_NOOWNERZORDER
dwFLAGS(7) = SWP_NOREDRAW
dwFLAGS(8) = SWP_NOREPOSITION
dwFLAGS(9) = SWP_NOSIZE
dwFLAGS(10) = SWP_NOZORDER
dwFLAGS(11) = SWP_SHOWWINDOW
dwFLAGS(12) = SWP_NOSENDCHANGING
dwFLAGS(13) = SWP_DEFERERASE

frmChildWindows.Show
IamActivated = False

End Sub

Private Sub cmdEnumWinProc_Click()

   lstEnumWindows.Clear

  'enumerate the windows passing the AddressOf the
  'callback function.  This example doesn't use the
  'lParam member.
   Call EnumWindows(AddressOf EnumWindowProc, &H0)

  'show the window count
   lblWindowsFound = lstEnumWindows.ListCount & " windows found."
If IamActivated Then
    lstEnumWindows.SetFocus
    lstEnumWindows.ListIndex = 0
End If
End Sub


Private Sub InitializeCombo()
With cboMessages
    .AddItem "WM_DESTROY"
    .AddItem "WM_ENABLE"
    .AddItem "WM_QUIT"
    .AddItem "WM_SETFOCUS"
    .AddItem "WM_CLOSE"
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmChildWindows
Me.WindowState = vbMinimized
Unload Me
End Sub

Private Sub lstEnumWindows_Click()

cmdSendMessage.Enabled = True
cmdSetTopMost.Enabled = True
cmdUnsetTopMost.Enabled = True

fr(4).Enabled = True
fr(8).Enabled = True

'Retrieve the hWnd of the selected window.
CurrenthWnd = lstEnumWindows.ItemData(lstEnumWindows.ListIndex)
OwnerhWnd = GetWindow(CurrenthWnd, GW_OWNER)

'Set the captions in the status bar.
lblCurrenthWnd.Caption = "Selected hWnd: " & str$(CurrenthWnd)
lblOwnerHwnd.Caption = ""
lblOwnerHwnd.Caption = "Owner hWnd: " & str$(OwnerhWnd)

'Refresh the Child Windows form
frmChildWindows.lstChildWindows.Clear
frmChildWindows.RefreshMe

End Sub

Private Sub optHWndInsertAfter_Click(Index As Integer)
Select Case Index
    Case 0:
        hWndInsertAfter = HWND_BOTTOM
        txthWndInsertAfter.Enabled = False
    Case 1:
        hWndInsertAfter = HWND_NOTOPMOST
        txthWndInsertAfter.Enabled = False
    Case 2:
        hWndInsertAfter = HWND_TOP
        txthWndInsertAfter.Enabled = False
    Case 3:
        hWndInsertAfter = HWND_TOPMOST
        txthWndInsertAfter.Enabled = False
    Case 4:
        txthWndInsertAfter.Enabled = True
        hWndInsertAfter = -666
End Select
        
End Sub

Private Sub txtHeight_GotFocus()
MakeSelection txtHeight
End Sub

Private Sub txthWndInsertAfter_GotFocus()
MakeSelection txthWndInsertAfter
End Sub

Private Sub txtNewWindowTitle_GotFocus()
MakeSelection txtNewWindowTitle
End Sub

Private Sub txtScreenLeft_GotFocus()
MakeSelection txtScreenLeft
End Sub

Private Sub txtScreenTop_GotFocus()
MakeSelection txtScreenTop
End Sub

Private Sub txtWidth_GotFocus()
MakeSelection txtWidth
End Sub

Private Sub txtWindowHeight_GotFocus()
MakeSelection txtWindowHeight
End Sub

Private Sub txtWindowWidth_GotFocus()
MakeSelection txtWindowWidth
End Sub
