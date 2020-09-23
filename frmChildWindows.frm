VERSION 5.00
Begin VB.Form frmChildWindows 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Child Windows"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstChildWindows 
      Height          =   4545
      ItemData        =   "frmChildWindows.frx":0000
      Left            =   0
      List            =   "frmChildWindows.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   8145
   End
   Begin VB.Frame fr 
      Height          =   525
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Width           =   3525
      Begin VB.TextBox txtNewText 
         Height          =   285
         Left            =   780
         TabIndex        =   6
         Top             =   150
         Width           =   1635
      End
      Begin VB.CommandButton cmdSetText 
         Caption         =   "Set Text"
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   150
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "New Text"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   7
         Top             =   210
         Width           =   690
      End
   End
   Begin VB.Frame fr 
      Height          =   525
      Index           =   2
      Left            =   3540
      TabIndex        =   8
      Top             =   4680
      Width           =   3345
      Begin VB.CommandButton cmdRightClick 
         Caption         =   "Right Click"
         Height          =   315
         Left            =   2250
         TabIndex        =   11
         Top             =   150
         Width           =   1005
      End
      Begin VB.CommandButton cmdLeftClick 
         Caption         =   "Left Click"
         Height          =   315
         Left            =   1185
         TabIndex        =   10
         Top             =   150
         Width           =   1005
      End
      Begin VB.CommandButton cmdDblClick 
         Caption         =   "Dbl Click"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   1005
      End
   End
   Begin VB.Frame fr 
      Height          =   525
      Index           =   1
      Left            =   6900
      TabIndex        =   12
      Top             =   4680
      Width           =   1245
      Begin VB.CommandButton cmdChangeStyle 
         Caption         =   "Change Style"
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   150
         Width           =   1125
      End
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Class"
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "hWnd"
      Height          =   195
      Index           =   1
      Left            =   2820
      TabIndex        =   2
      Top             =   0
      Width           =   435
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Window Text"
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   1
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmChildWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Local variables
Private CurrentStyle As Long
Private ButtonStyles(18) As Long
Private MaxStyles As Byte

'Change the style of the selected button
'If the window is not a button, then the message is disregarded.
'Will loop through all the available styles.
Private Sub cmdChangeStyle_Click()

Call SendMessage(Childhwnd, BM_SETSTYLE, CurrentStyle, 1)
CurrentStyle = (CurrentStyle + 1) Mod (MaxStyles + 1)
End Sub

'Send a Double - Click event
Private Sub cmdDblClick_Click()

Call SetActiveWindow(CurrenthWnd)
Call SendMessage(Childhwnd, WM_LBUTTONDBLCLK, MK_LBUTTON, 0&)
Call SendMessage(Childhwnd, BN_DOUBLECLICKED, 0&, 0&)
End Sub

'Send a simple left click event.
Private Sub cmdLeftClick_Click()

Call SendMessage(Childhwnd, BM_CLICK, 0&, 0&)
End Sub

'Send a right click event
Private Sub cmdRightClick_Click()

Call SetActiveWindow(CurrenthWnd)
Call SendMessage(Childhwnd, WM_RBUTTONDOWN, MK_RBUTTON, 0&)
Call SendMessage(Childhwnd, WM_RBUTTONUP, MK_RBUTTON, 0&)
End Sub

'Change the window's text.
Private Sub cmdSetText_Click()
Dim dummy() As Byte

ReDim dummy(Len(txtNewText) + 1)

'The string has to be converted to a byte array to be passed to the function
'I guess it could be done otherwise, but it works this way also.
dummy = StringToByteArray(txtNewText.Text)

'dummy(0) points to the first byte of the array
Call SendMessage(Childhwnd, WM_SETTEXT, 0&, dummy(0))

End Sub

Private Sub Form_Load()
Dim TabArray(0 To 2) As Long
   
TabArray(0) = 0
TabArray(1) = -142
TabArray(2) = 154
   
'clear any existing tabs
Call SendMessage(lstChildWindows.hWnd, LB_SETTABSTOPS, 0&, ByVal 0&)
   
'set list tabstops
Call SendMessage(lstChildWindows.hWnd, LB_SETTABSTOPS, 3&, TabArray(0))

'Initialize the array where the button styles are stored.
ButtonStyles(0) = BS_PUSHLIKE
ButtonStyles(1) = BS_RADIOBUTTON
ButtonStyles(2) = BS_AUTORADIOBUTTON
ButtonStyles(3) = BS_CHECKBOX
ButtonStyles(4) = BS_3STATE
ButtonStyles(5) = BS_AUTO3STATE
ButtonStyles(6) = BS_AUTOCHECKBOX
ButtonStyles(7) = BS_PUSHBUTTON
ButtonStyles(8) = BS_GROUPBOX
ButtonStyles(9) = BS_DEFPUSHBUTTON
ButtonStyles(10) = BS_LEFTTEXT
ButtonStyles(11) = BS_SOLID
ButtonStyles(12) = BS_BOTTOM
ButtonStyles(13) = BS_CENTER
ButtonStyles(14) = BS_LEFT
ButtonStyles(15) = BS_MULTILINE
ButtonStyles(16) = BS_RIGHT
ButtonStyles(17) = BS_TOP
ButtonStyles(18) = BS_VCENTER

'Initialize the current style.
CurrentStyle = ButtonStyles(0)
MaxStyles = UBound(ButtonStyles)
frmMain.optWindow(1).Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.optWindow(1).Enabled = False
frmMain.optWindow(0).Value = True
End Sub

Private Sub lstChildWindows_Click()
Childhwnd = lstChildWindows.ItemData(lstChildWindows.ListIndex)
End Sub
Public Sub RefreshMe()
Call EnumChildWindows(CurrenthWnd, AddressOf EnumChildWindowProc, 0&)
Childhwnd = 0
If lstChildWindows.ListCount > 0 Then
    lstChildWindows.ListIndex = 0
    frmMain.optWindow(1).Enabled = True
Else
    frmMain.optWindow(1).Enabled = False
    frmMain.optWindow(0).Value = True
End If
End Sub
