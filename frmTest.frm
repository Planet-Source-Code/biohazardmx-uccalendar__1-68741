VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "ucCalendar Demonstration"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   451
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   150
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   235
      TabIndex        =   15
      Top             =   3270
      Width           =   3525
      Begin VB.CheckBox Check4 
         Caption         =   "Output status messages"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   1860
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2250
         TabIndex        =   18
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   60
         TabIndex        =   17
         Top             =   330
         Width           =   3405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Event Log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   45
         TabIndex        =   16
         Top             =   15
         Width           =   825
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000010&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   3795
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   3810
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   1
      Top             =   180
      Width           =   2805
      Begin VB.CheckBox Check5 
         Caption         =   "Show week numbers"
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   3690
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmTest.frx":0000
         Left            =   60
         List            =   "frmTest.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2430
         Width           =   2685
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show month change buttons"
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   3420
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show week day names"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   3150
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show month name and year"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmTest.frx":0076
         Left            =   60
         List            =   "frmTest.frx":0083
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1800
         Width           =   2685
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmTest.frx":00AE
         Left            =   60
         List            =   "frmTest.frx":00BE
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1170
         Width           =   2685
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmTest.frx":00ED
         Left            =   60
         List            =   "frmTest.frx":0106
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   540
         Width           =   2685
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "First day of week:"
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   2220
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Focus style:"
         Height          =   195
         Left            =   60
         TabIndex        =   8
         Top             =   1590
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Line style:"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Border style:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Properties"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   45
         TabIndex        =   2
         Top             =   15
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000010&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   2805
      End
   End
   Begin ucCalendarTest.ucCalendar ucCalendar1 
      Height          =   2700
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   4763
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ActiveDayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MonthNameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderBackColor =   -2147483633
      HeaderForeColor =   -2147483630
      ShowWeekNumber  =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   225
      Left            =   150
      TabIndex        =   9
      Top             =   2970
      Width           =   3525
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LoadLibrary Lib "Kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Sub InitCommonControls Lib "ComCtl32.dll" ()

Private hShellLib As Long

Private Sub Check1_Click()
  ucCalendar1.ShowMonth = CBool(Check1.Value)
End Sub
Private Sub Check2_Click()
  ucCalendar1.ShowWeekDays = CBool(Check2.Value)
End Sub
Private Sub Check3_Click()
  ucCalendar1.MonthChangeButtons = CBool(Check3.Value)
End Sub
Private Sub Check5_Click()
  ucCalendar1.ShowWeekNumber = CBool(Check5.Value)
End Sub

Private Sub Combo1_Click()
  ucCalendar1.BorderStyle = Combo1.ListIndex
End Sub
Private Sub Combo2_Click()
  ucCalendar1.LineStyle = Combo2.ListIndex
End Sub
Private Sub Combo3_Click()
  ucCalendar1.FocusStyle = Combo3.ListIndex
End Sub
Private Sub Combo4_Click()
  ucCalendar1.FirstDayOfWeek = Combo4.ListIndex
End Sub

Private Sub Command1_Click()
  Call List1.Clear
End Sub

Private Sub Form_Initialize()
  hShellLib = LoadLibrary("Shell32.dll")
  Call InitCommonControls
End Sub
Private Sub Form_Load()
  Caption = "ucCalendar " & ucCalendar1.ControlVersion & " - Control Demonstration"
  Combo1.ListIndex = ucCalendar1.BorderStyle
  Combo2.ListIndex = ucCalendar1.LineStyle
  Combo3.ListIndex = ucCalendar1.FocusStyle
  Combo4.ListIndex = ucCalendar1.FirstDayOfWeek
  Label5.Caption = Format(ucCalendar1.ActiveDate, "Long Date")
  Call pvLogEvent("Welcome, using ucCalendar " & ucCalendar1.ControlVersion)
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Call FreeLibrary(hShellLib)
End Sub

Private Sub pvLogEvent(ByVal sText As String)
  If (Check4.Value = vbUnchecked) Then Exit Sub
  Call List1.AddItem(sText)
  List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub ucCalendar1_Change()
  Label5.Caption = Format(ucCalendar1.ActiveDate, "Long Date")
  Call pvLogEvent("Date changed to " & Format(ucCalendar1.ActiveDate, "Short Date"))
End Sub
