VERSION 5.00
Begin VB.UserControl ucCalendar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "ucCalendar.ctx":0000
   PropertyPages   =   "ucCalendar.ctx":985E
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ToolboxBitmap   =   "ucCalendar.ctx":9881
   Begin VB.Menu mnuMonths 
      Caption         =   "Months"
      Begin VB.Menu mnuMonth 
         Caption         =   "January"
         Index           =   1
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "Febraury"
         Index           =   2
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "March"
         Index           =   3
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "April"
         Index           =   4
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "May"
         Index           =   5
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "June"
         Index           =   6
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "July"
         Index           =   7
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "August"
         Index           =   8
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "September"
         Index           =   9
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "October"
         Index           =   10
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "November"
         Index           =   11
      End
      Begin VB.Menu mnuMonth 
         Caption         =   "December"
         Index           =   12
      End
   End
   Begin VB.Menu mnuYears 
      Caption         =   "Years"
      Begin VB.Menu mnuYear 
         Caption         =   "-5"
         Index           =   1
      End
      Begin VB.Menu mnuYear 
         Caption         =   "-4"
         Index           =   2
      End
      Begin VB.Menu mnuYear 
         Caption         =   "-3"
         Index           =   3
      End
      Begin VB.Menu mnuYear 
         Caption         =   "-2"
         Index           =   4
      End
      Begin VB.Menu mnuYear 
         Caption         =   "-1"
         Index           =   5
      End
      Begin VB.Menu mnuYear 
         Caption         =   "Current"
         Checked         =   -1  'True
         Index           =   6
      End
      Begin VB.Menu mnuYear 
         Caption         =   "+1"
         Index           =   7
      End
      Begin VB.Menu mnuYear 
         Caption         =   "+2"
         Index           =   8
      End
      Begin VB.Menu mnuYear 
         Caption         =   "+3"
         Index           =   9
      End
      Begin VB.Menu mnuYear 
         Caption         =   "+4"
         Index           =   10
      End
      Begin VB.Menu mnuYear 
         Caption         =   "+5"
         Index           =   11
      End
   End
End
Attribute VB_Name = "ucCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************
' Name:      ucCalendar
' Version:   0.9 RC 1
' Date:      25/05/2007
' Author:    BioHazardMX
' Details:   Calendar Control (self-contained, full API drawn)
' Requires:  Win32 Libraries (User32, GDI32, Kernel32)
'***************************************************************************************
' Changes
'
' Version 0.9 RC 1 (1,513 lines)
'  * New toolbox icon
'  * Release Candidate 1, the next version may be a final release
'  * Added descriptions for procedures, properties and events (for Properties window)
'  * Added 'pvGetCellByDay' and 'pvGetDayByCell' procedures (simplifying other procedures)
'  * Fixed small issue with months that start in Sunday (or the specified first week day)
'  * Fixed automatic month changer, the last day will be clipped accordingly to each month
'  * Fixed memory leak in pvDrawControl (not deleting the symbol font due to a mistype)
'  * Updated 'pvGetPeriodBounds', It'll only search the period bounds when month has changed
'  * Added SelfSub-powered subclassing to detect theme and color changes to redraw control
'  * Updated month change menu, now it uses Office XP style under WinXP when themes are activated
'  * Fixed: ShadedRect + SimpleLine = Overlapped borders in the 7th column's right border
'  * Added 'ShowWeekNumber' property, toggle week numbers on the left of the calendar
'  * Fixed resizing behaviour, finally I've found a working (at least for me) workaround
'  * Added a new adaptive year change menu, it shows the previous/next 5 years to choose from!
'  * Fixed menus appearing in MouseDown event (now they appear with MouseUp)
'  * Fixed stupid bug with month change buttons even when the month name bar where not visible!
'
' Version 0.8 Alpha (602 lines)
'  * First alpha version, almost feature-complete
'  * Added autosizing feature, now the control is adjusted to show only complete cells
'  * Fixed issue when changing border style, now the property fires the Resize event
'  * Added support for non-sunday-starting weeks (through 'FirstDayOfWeek' property)
'  * Added 'ShowMonth' and 'ShowDayNames' properties
'  * Fixed small issue in pvLocationToCell parameters (they where inverted)
'  * Updated drawing procedures, now using an off-screen DC for faster rendering
'  * Added a basic keyboard interface
'  * Added automatic month change when clicking previous/next month days
'  * Added month change menu (click menu name)
'  * Added month change buttons (optional, see 'MonthChangeButtons' property )
'  * Added 'FocusStyle' property to customize the selection rectangle style
'  * Fixed small glitch when drawing after receiving focus with mouse down
'  * Code rearranged, unused stuff was removed and procedures have been simplified
'
'***************************************************************************************
' ToDo
'
'  * Maybe year change buttons (with the adaptive menu this isn't required)
'  * Maybe visual themes (Win9x, WinXP, OfficeXP, Office 2003, MacOSX, etc)
'  * Highligthing of specific dates (not only the active) to show several events in one
'    calendar (like Rainlendar does)
'  * More MonthView functionality (week numbers[DONE], multiple selections, etc)
'
'  Can you help me? Have you already added one of these features?
'    If so, please contact me to update the control.
'
'***************************************************************************************
' License
'
'  You may use this control and/or any part of its source code freely as long as you
'  agree that no warranty or liability of any kind is expressed or implied. You may not
'  remove any copyright notices and also you may not claim ownership unless you have
'  substantially altered the functionality or features.
'
'  This software uses the SelfSub subclassing system from Paul Caton.
'  Thanks to Paul Caton for this excellent code, find out more at:
'   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'  Copyright © 2007 BioHazardMX.
'  Proudly made in Mexico.
'
'***************************************************************************************
 Option Explicit
'***************************************************************************************
' Constants
'***************************************************************************************
'---Control Specific
Private Const CTL_VERSION As String = "0.9 RC1"  'PLEASE DON'T MODIFY UNLESS YOU MAKE CONSISTENT CHANGES
Private Const CTL_IDE_SUBCLASS As Boolean = True 'Toggle IDE Subclassing On/Off
'---DrawText Flags
Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER As Long = &H4
Private Const DT_WORD_ELLIPSIS As Long = &H40000
'---OwnerDrawn States
Private Const ODS_CHECKED As Long = &H8
Private Const ODS_GRAYED As Long = &H2
Private Const ODS_SELECTED As Long = &H1
'---Messages
Private Const WM_THEMECHANGED As Long = &H31A
Private Const WM_SYSCOLORCHANGE As Long = &H15 '21
Private Const WM_INITMENUPOPUP As Long = &H117
Private Const WM_DRAWITEM As Long = &H2B
Private Const WM_MEASUREITEM As Long = &H2C
'***************************************************************************************
' UDT's
'***************************************************************************************
'---RECT
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
'---MENUITEMINFO
Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type
'---DRAWITEMSTRUCT
Private Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hWndItem As Long
  hDC As Long
  rcItem As RECT
  itemData As Long
End Type
'---MEASUREITEMSTRUCT
Private Type MEASUREITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemWidth As Long
  itemHeight As Long
  itemData As Long
End Type
'---LOGFONT
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(1 To 32) As Byte
End Type
'---NONCLIENTMETRICS
Private Type NONCLIENTMETRICS
  cbSize As Long
  iBorderWidth As Long
  iScrollWidth As Long
  iScrollHeight As Long
  iCaptionWidth As Long
  iCaptionHeight As Long
  lfCaptionFont As LOGFONT
  iSMCaptionWidth As Long
  iSMCaptionHeight As Long
  lfSMCaptionFont As LOGFONT
  iMenuWidth As Long
  iMenuHeight As Long
  lfMenuFont As LOGFONT
  lfStatusFont As LOGFONT
  lfMessageFont As LOGFONT
End Type
'***************************************************************************************
' API Declares
'***************************************************************************************
'---Kernel32
Private Declare Function MulDiv Lib "Kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function FreeLibrary Lib "Kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "Kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'---User32
Private Declare Function GetSysColor Lib "User32.dll" (ByVal nIndex As Long) As Long
Private Declare Function DrawEdge Lib "User32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal uEdge As Long, ByVal uFlags As Long) As Long
Private Declare Function FillRect Lib "User32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "User32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "User32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PtInRect Lib "User32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawText Lib "User32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "User32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetMenuItemInfo Lib "User32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal uByPosition As Long, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "User32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal uByPosition As Long, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemCount Lib "User32.dll" (ByVal hMenu As Long) As Long
Private Declare Function SystemParametersInfo Lib "User32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
'---GDI32
Private Declare Function BitBlt Lib "GDI32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateDCAsNull Lib "GDI32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateFontIndirect Lib "GDI32.dll" Alias "CreateFontIndirectA" (ByRef lpLogFont As LOGFONT) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateFontA Lib "GDI32.dll" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function CreateSolidBrush Lib "GDI32.dll" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "GDI32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "GDI32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function DeleteObject Lib "GDI32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32.dll" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "GDI32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
'---UxTheme (only WinXP+)
Private Declare Function OpenThemeData Lib "UxTheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme.dll" (ByVal hTheme As Long) As Long
'***************************************************************************************
' SelfSub Code
'***************************************************************************************
'---Enumerations
Private Enum eMsgWhen
  MSG_BEFORE = 1
  MSG_AFTER = 2
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum
'---Constants
Private Const ALL_MESSAGES  As Long = -1
Private Const MSG_ENTRIES   As Long = 32
Private Const WNDPROC_OFF   As Long = &H38
Private Const GWL_WNDPROC   As Long = -4
Private Const IDX_SHUTDOWN  As Long = 1
Private Const IDX_HWND      As Long = 2
Private Const IDX_WNDPROC   As Long = 9
Private Const IDX_BTABLE    As Long = 11
Private Const IDX_ATABLE    As Long = 12
Private Const IDX_PARM_USER As Long = 13
'---API Declares
Private Declare Function CallWindowProcA Lib "User32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "Kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "Kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "Kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "Kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "Kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'---Properties
Private z_ScMem As Long
Private z_Sc(64) As Long
Private z_Funk As Collection
'***************************************************************************************
' Enumerations
'***************************************************************************************
'---ucCalendarLineStyles
Public Enum ucCalendarLineStyles
  [clsNoLine] = 0
  [clsRaised]
  [clsSunken]
  [clsSimple]
End Enum
'---ucCalendarBorderStyles
Public Enum ucCalendarBorderStyles
  [cbsNone] = 0
  [cbsSunken]
  [cbsRaised]
  [cbsFlatSunken]
  [cbsFlatRaised]
  [cbsEtched]
  [cbsSimple]
End Enum
'---ucCalendarFocusStyles
Public Enum ucCalendarFocusStyles
  [cfsNone] = 0
  [cfsFocusRect]
  [cfsShadedRect]
End Enum
'***************************************************************************************
' Variables
'***************************************************************************************
Private bIsLeapYear As Boolean
Private bPropLoaded As Boolean
Private bShowMonth As Boolean
Private bShowWeekDays As Boolean
Private bShowWeekNum As Boolean
Private bMonthButtons As Boolean
Private bRunning As Boolean
Private bHasFocus As Boolean
Private bMouseDown As Boolean
Private bIgnoreEvent As Boolean
Private sActiveDate As String
Private lHeaderBack As Long
Private lHeaderFore As Long
Private lDaysBack As Long
Private lDaysFore As Long
Private lOutDaysBack As Long
Private lOutDaysFore As Long
Private lActiveDayFore As Long
Private lSelectBack As Long
Private lSelectFore As Long
Private lLineColor As Long
Private lBorderColor As Long
Private lLineStyle As Long
Private lBorderStyle As Long
Private lFocusStyle As Long
Private lFirstWeekDay As Long
Private lActiveCell As Long
Private lYear As Long
Private lMonth As Long
Private lDay As Long
Private tSelFont As StdFont
Private tHdrFont As StdFont
Private sCaptions(25) As String
'***************************************************************************************
' Events
'***************************************************************************************
Public Event Change()
Attribute Change.VB_Description = "Raised whenever the user changes the currently selected date."
Public Event Click()
Attribute Click.VB_Description = "Raised when the mouse is pressed and then released over the control."
'***************************************************************************************
' Properties
'***************************************************************************************
'---ControlVersion:
Public Property Get ControlVersion() As String
Attribute ControlVersion.VB_Description = "Returns the version number information for this control."
  ControlVersion = CTL_VERSION
End Property
'---Enabled: <Description>
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Determines whether an object responds to user-generated events or not."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
Attribute Enabled.VB_UserMemId = -514
  Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal vData As Boolean)
  UserControl.Enabled = vData
  Call PropertyChanged("Enabled")
  Call pvDrawControl
End Property
'---ShowMonth: <Description>
Public Property Get ShowMonth() As Boolean
Attribute ShowMonth.VB_Description = "Determines whether to show the month name and year or not."
Attribute ShowMonth.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
  ShowMonth = bShowMonth
End Property
Public Property Let ShowMonth(ByVal vData As Boolean)
Dim lHeight As Long
  bShowMonth = vData
  Call PropertyChanged("ShowMonth")
  Call UserControl_Resize
  If (bShowMonth) Then lHeight = Height + (19 * Screen.TwipsPerPixelX) Else lHeight = Height - (19 * Screen.TwipsPerPixelX)
  Height = lHeight
End Property
'---ShowWeekDays: <Description>
Public Property Get ShowWeekDays() As Boolean
Attribute ShowWeekDays.VB_Description = "Determines whether to show the week day names or not."
Attribute ShowWeekDays.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
  ShowWeekDays = bShowWeekDays
End Property
Public Property Let ShowWeekDays(ByVal vData As Boolean)
Dim lHeight As Long
  bShowWeekDays = vData
  Call PropertyChanged("ShowWeekDays")
  Call UserControl_Resize
  If (bShowWeekDays) Then lHeight = Height + (19 * Screen.TwipsPerPixelX) Else lHeight = Height - (19 * Screen.TwipsPerPixelX)
  Height = lHeight
End Property
'---MonthChangeButtons: <Description>
Public Property Get MonthChangeButtons() As Boolean
Attribute MonthChangeButtons.VB_Description = "Determines whether the control has month change buttons or not."
  MonthChangeButtons = bMonthButtons
End Property
Public Property Let MonthChangeButtons(ByVal vData As Boolean)
  bMonthButtons = vData
  Call PropertyChanged("MonthChangeButtons")
  If (bShowMonth) Then Call pvDrawControl
End Property
'---ShowWeekNumber: <Description>
Public Property Get ShowWeekNumber() As Boolean
Attribute ShowWeekNumber.VB_Description = "Determines whether to show the week numbers or not."
Attribute ShowWeekNumber.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
  ShowWeekNumber = bShowWeekNum
End Property
Public Property Let ShowWeekNumber(ByVal vData As Boolean)
Dim lWidth As Long
  bShowWeekNum = vData
  Call PropertyChanged("ShowWeekNumber")
  If (bShowWeekNum) Then lWidth = Width + (24 * Screen.TwipsPerPixelX) Else lWidth = Width - (24 * Screen.TwipsPerPixelX)
  Width = lWidth
End Property
'---ActiveDate: <Description>
Public Property Get ActiveDate() As String
Attribute ActiveDate.VB_Description = "Returns or sets a value that indicates the currently selected date."
Attribute ActiveDate.VB_ProcData.VB_Invoke_Property = ";Otras"
Attribute ActiveDate.VB_UserMemId = 0
  ActiveDate = sActiveDate
End Property
Public Property Let ActiveDate(ByVal vData As String)
  On Error GoTo ProcExit
  sActiveDate = DateValue(vData)
  Call PropertyChanged("ActiveDate")
  lYear = DatePart("yyyy", sActiveDate)
  lMonth = DatePart("m", sActiveDate)
  lDay = DatePart("d", sActiveDate)
  Call pvUpdateYearsMenu
  Call pvCheckMonthMenu
  Call pvDrawControl
ProcExit:
End Property
'---LineStyle: <Description>
Public Property Get LineStyle() As ucCalendarLineStyles
Attribute LineStyle.VB_Description = "Returns or sets the style that will be used to draw lines between cells."
Attribute LineStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  LineStyle = lLineStyle
End Property
Public Property Let LineStyle(ByVal vData As ucCalendarLineStyles)
  lLineStyle = vData
  Call PropertyChanged("LineStyle")
  Call pvDrawControl
End Property
'---BorderStyle: <Description>
Public Property Get BorderStyle() As ucCalendarBorderStyles
Attribute BorderStyle.VB_Description = "Returns or sets the style that will be used to draw the control border."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  BorderStyle = lBorderStyle
End Property
Public Property Let BorderStyle(ByVal vData As ucCalendarBorderStyles)
Dim lBorderW As Long
Dim lPrevW As Long
Dim lOffset As Long
  Select Case lBorderStyle
    Case 0: lPrevW = 0
    Case 1, 2, 5: lPrevW = 2
    Case 3, 4, 6: lPrevW = 1
  End Select
  lBorderStyle = vData
  Call PropertyChanged("BorderStyle")
  Select Case lBorderStyle
    Case 0: lBorderW = 0
    Case 1, 2, 5: lBorderW = 2
    Case 3, 4, 6: lBorderW = 1
  End Select
  lOffset = (lPrevW - lBorderW) * -1
  Width = Width + lOffset
  Height = Height + lOffset
  Call UserControl_Resize
End Property
'---FocusStyle: <Description>
Public Property Get FocusStyle() As ucCalendarFocusStyles
Attribute FocusStyle.VB_Description = "Returns or sets the style that will be used to draw the  selection rectangle when the control has the focus."
Attribute FocusStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  FocusStyle = lFocusStyle
End Property
Public Property Let FocusStyle(ByVal vData As ucCalendarFocusStyles)
  lFocusStyle = vData
  Call PropertyChanged("FocusStyle")
  If bRunning Then Call pvDrawControl
End Property
'---FirstDayOfWeek: <Description>
Public Property Get FirstDayOfWeek() As VbDayOfWeek
Attribute FirstDayOfWeek.VB_Description = "Returns or sets the day that will be used as the first day of the week."
Attribute FirstDayOfWeek.VB_ProcData.VB_Invoke_Property = ";Comportamiento"
  FirstDayOfWeek = lFirstWeekDay
End Property
Public Property Let FirstDayOfWeek(ByVal vData As VbDayOfWeek)
  lFirstWeekDay = vData
  Call PropertyChanged("FirstDayOfWeek")
  Call pvDrawControl
End Property
'---LineColor: <Description>
Public Property Get LineColor() As OLE_COLOR
Attribute LineColor.VB_Description = "Returns or sets the color that will be used to draw the lines between cells (in simple line mode)."
Attribute LineColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  LineColor = lLineColor
End Property
Public Property Let LineColor(ByVal vData As OLE_COLOR)
  lLineColor = vData
  Call PropertyChanged("LineColor")
  Call pvDrawControl
End Property
'---BorderColor: <Description>
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns or sets the color of the control border (for simple border style)."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  BorderColor = lBorderColor
End Property
Public Property Let BorderColor(ByVal vData As OLE_COLOR)
  lBorderColor = vData
  Call PropertyChanged("BorderColor")
  Call pvDrawControl
End Property
'---HeaderBackColor: <Description>
Public Property Get HeaderBackColor() As OLE_COLOR
Attribute HeaderBackColor.VB_Description = "Returns or sets the background color for the month and week day names."
Attribute HeaderBackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  HeaderBackColor = lHeaderBack
End Property
Public Property Let HeaderBackColor(ByVal vData As OLE_COLOR)
  lHeaderBack = vData
  Call PropertyChanged("HeaderBackColor")
  Call pvDrawControl
End Property
'---HeaderForeColor: <Description>
Public Property Get HeaderForeColor() As OLE_COLOR
Attribute HeaderForeColor.VB_Description = "Returns or sets the text color for the month and week day names."
Attribute HeaderForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  HeaderForeColor = lHeaderFore
End Property
Public Property Let HeaderForeColor(ByVal vData As OLE_COLOR)
  lHeaderFore = vData
  Call PropertyChanged("HeaderForeColor")
  Call pvDrawControl
End Property
'---DaysBackColor: <Description>
Public Property Get DaysBackColor() As OLE_COLOR
Attribute DaysBackColor.VB_Description = "Returns or sets the background color for the current month days."
Attribute DaysBackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  DaysBackColor = lDaysBack
End Property
Public Property Let DaysBackColor(ByVal vData As OLE_COLOR)
  lDaysBack = vData
  Call PropertyChanged("DaysBackColor")
  Call pvDrawControl
End Property
'---DaysForeColor: <Description>
Public Property Get DaysForeColor() As OLE_COLOR
Attribute DaysForeColor.VB_Description = "Returns or sets the text color for the active current days."
Attribute DaysForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  DaysForeColor = lDaysFore
End Property
Public Property Let DaysForeColor(ByVal vData As OLE_COLOR)
  lDaysFore = vData
  Call PropertyChanged("DaysForeColor")
  Call pvDrawControl
End Property
'---OutDaysBackColor: <Description>
Public Property Get OutDaysBackColor() As OLE_COLOR
Attribute OutDaysBackColor.VB_Description = "Returns or sets the background color for the previous/next month days."
Attribute OutDaysBackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  OutDaysBackColor = lOutDaysBack
End Property
Public Property Let OutDaysBackColor(ByVal vData As OLE_COLOR)
  lOutDaysBack = vData
  Call PropertyChanged("OutDaysBackColor")
  Call pvDrawControl
End Property
'---OutDaysForeColor: <Description>
Public Property Get OutDaysForeColor() As OLE_COLOR
Attribute OutDaysForeColor.VB_Description = "Returns or sets the text color for the previous/next month days."
Attribute OutDaysForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  OutDaysForeColor = lOutDaysFore
End Property
Public Property Let OutDaysForeColor(ByVal vData As OLE_COLOR)
  lOutDaysFore = vData
  Call PropertyChanged("OutDaysForeColor")
  Call pvDrawControl
End Property
'---ActiveDayForeColor: <Description>
Public Property Get ActiveDayForeColor() As OLE_COLOR
Attribute ActiveDayForeColor.VB_Description = "Returns or sets the color of the currently selected day number."
Attribute ActiveDayForeColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  ActiveDayForeColor = lActiveDayFore
End Property
Public Property Let ActiveDayForeColor(ByVal vData As OLE_COLOR)
  lActiveDayFore = vData
  Call PropertyChanged("ActiveDayForeColor")
  Call pvDrawControl
End Property
'---ActiveDayFont: <Description>
Public Property Get ActiveDayFont() As StdFont
Attribute ActiveDayFont.VB_Description = "Returns or sets a font object that identifies the font used to draw the active day number."
Attribute ActiveDayFont.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  Set ActiveDayFont = tSelFont
End Property
Public Property Set ActiveDayFont(ByVal vData As StdFont)
  Set tSelFont = vData
  Call PropertyChanged("ActiveDayFont")
  Call pvDrawControl
End Property
'---MonthNameFont: <Description>
Public Property Get MonthNameFont() As StdFont
Attribute MonthNameFont.VB_Description = "Returns or sets a font object that identifies the font used to draw the month name and year."
Attribute MonthNameFont.VB_ProcData.VB_Invoke_Property = ";Apariencia"
  Set MonthNameFont = tHdrFont
End Property
Public Property Set MonthNameFont(ByVal vData As StdFont)
  Set tHdrFont = vData
  Call PropertyChanged("MonthNameFont")
  Call pvDrawControl
End Property
'---Font: <Description>
Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns or sets a font object that identifies the font used to draw text on the control."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal vData As StdFont)
  Set UserControl.Font = vData
  Call PropertyChanged("Font")
  Call pvDrawControl
End Property
'***************************************************************************************
' Procedures
'***************************************************************************************
'---AboutBox
Public Sub AboutBox()
Attribute AboutBox.VB_Description = "Show copyright and version number information for this control."
Attribute AboutBox.VB_UserMemId = -552
  Call MsgBox("ucCalendar Control " & CTL_VERSION & vbCrLf & vbCrLf & "A simple but powerful, self-contained calendar control with" & vbCrLf & "customizable appeareance and standard functionality." & vbCrLf & vbCrLf & "Written by BioHazardMX" & vbCrLf & "Copyright © 2007 BioHazardMX" & vbCrLf & vbCrLf & "You may use this control and/or any part of its source code freely" & vbCrLf & "as long as you agree that no warranty or liability of any kind is" & vbCrLf & "expressed or implied. You may not remove any copyright notices" & vbCrLf & "and also you may not claim ownership unless you have substantially" & vbCrLf & "altered the functionality or features.", vbOKOnly, "About ucCalendar")
End Sub
'---Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Force a control or an object to repaint itself."
Attribute Refresh.VB_UserMemId = -550
  Call pvDrawControl
End Sub
'---pvDrawControl:
Private Sub pvDrawControl()
Dim rcRect As RECT
Dim rcButton(1) As RECT
Dim rcCell(54) As RECT
Dim lX As Long, lY As Long
Dim lCellW As Long, lCellH As Long
Dim lCell As Long, sBuffer As String
Dim hFont As Long, lFont As Long
Dim hSelFont As Long, lSelFont As Long
Dim hHdrFont As Long, lHdrFont As Long
Dim hSymFont As Long, lSymFont As Long
Dim hBrush As Long, bActive As Boolean
Dim lFirstCell As Long, lLastCell As Long
Dim lLastMonthDay As Long, lOffset As Long
Dim lBorderW As Long, lHeaderH As Long
Dim lhDC As Long, lhBmp As Long, loBmp As Long
Dim lRet As Long, lColor As Long
 'If the properties haven't been loaded exit
  If (Not bPropLoaded) Then Exit Sub
 'Create an off-screen canvas and the logic fonts
  Call pvCreate(lhDC, lhBmp, loBmp, ScaleWidth, ScaleHeight)
  hFont = CreateFont(Font.Name, Font.Size, Font.Charset, Font.Bold, Font.Italic, Font.Underline, Font.Strikethrough)
  hSelFont = CreateFont(tSelFont.Name, tSelFont.Size, tSelFont.Charset, tSelFont.Bold, tSelFont.Italic, tSelFont.Underline, tSelFont.Strikethrough)
  hHdrFont = CreateFont(tHdrFont.Name, tHdrFont.Size, tHdrFont.Charset, tHdrFont.Bold, tHdrFont.Italic, tHdrFont.Underline, tHdrFont.Strikethrough)
  hSymFont = CreateFont("Marlett", 10)
 'Select the first font and set text mode to tranaparent
  lFont = SelectObject(lhDC, hFont)
  lRet = SetBkMode(lhDC, 1)
 'Initialize variables
  lLastMonthDay = DateSerial(lYear, lMonth + 1, 1 - 1)
  lLastMonthDay = DatePart("d", lLastMonthDay, lFirstWeekDay)
  Select Case lBorderStyle
    Case 0: lBorderW = 0
    Case 1, 2, 5: lBorderW = 2
    Case 3, 4, 6: lBorderW = 1
  End Select
  If (bShowMonth) Then lHeaderH = 19
  If (bShowWeekDays) Then lHeaderH = lHeaderH + 19
  If (bShowWeekNum) Then lOffset = 24
  Call pvGetPeriodBounds(lFirstCell, lLastCell)
  lCellW = (ScaleWidth - ((lBorderW * 2) + lOffset)) \ 7
  lCellH = (ScaleHeight - ((lBorderW * 2) + lHeaderH)) \ 6
  lCell = 7
 'Draw day cells
  For lY = 1 To 6
    For lX = 0 To 6
      sBuffer = pvGetDayByCell(lCell)
      bActive = False
      If Not pvIsWithinPeriod(lCell) Then
        lRet = SetTextColor(lhDC, pvGetColor(lOutDaysFore))
        hBrush = CreateSolidBrush(pvGetColor(lOutDaysBack))
      Else
        lRet = SetTextColor(lhDC, pvEvaluate(sBuffer <> lDay, pvGetColor(lDaysFore), pvGetColor(lActiveDayFore)))
        bActive = (sBuffer = lDay)
        hBrush = CreateSolidBrush(pvGetColor(lDaysBack))
      End If
      lRet = SetRect(rcCell(lCell), (lBorderW + lOffset) + (lX * lCellW), (lBorderW + lHeaderH) + ((lY - 1) * lCellH), (lBorderW + lOffset) + ((lX + 1) * lCellW), (lBorderW + lHeaderH) + (lY * lCellH))
      lRet = FillRect(lhDC, rcCell(lCell), hBrush)
      lRet = DeleteObject(hBrush)
      If (Not bActive) Then
        lRet = DrawText(lhDC, sBuffer, Len(sBuffer), rcCell(lCell), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
      Else
        If (lLineStyle = 0) Then
          lRet = SetRect(rcRect, rcCell(lCell).Left + 1, rcCell(lCell).Top + 1, rcCell(lCell).Right - 1, rcCell(lCell).Bottom - 1)
        ElseIf (lLineStyle = 3) Then
          lRet = SetRect(rcRect, rcCell(lCell).Left + 2, rcCell(lCell).Top + 2, rcCell(lCell).Right - pvEvaluate(lX = 6, 2, 1), rcCell(lCell).Bottom - 1)
        Else
          lRet = SetRect(rcRect, rcCell(lCell).Left + 1, rcCell(lCell).Top + 1, rcCell(lCell).Right - 2, rcCell(lCell).Bottom - 2)
        End If
        If (bHasFocus) Then
          Select Case lFocusStyle
            Case 1
              lColor = SetTextColor(lhDC, &H0)
              lRet = DrawFocusRect(lhDC, rcRect)
              lRet = SetTextColor(lhDC, lColor)
            Case 2
              hBrush = CreateSolidBrush(pvBlendColor(lActiveDayFore, lDaysBack, 50))
              lRet = FillRect(lhDC, rcRect, hBrush)
              lRet = DeleteObject(hBrush)
              hBrush = CreateSolidBrush(pvGetColor(lActiveDayFore))
              lRet = FrameRect(lhDC, rcRect, hBrush)
              lRet = DeleteObject(hBrush)
          End Select
        End If
        lSelFont = SelectObject(lhDC, hSelFont)
        lRet = DrawText(lhDC, sBuffer, Len(sBuffer), rcCell(lCell), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
        lRet = SelectObject(lhDC, lSelFont)
      End If
      Select Case lLineStyle
        Case 1: lRet = DrawEdge(lhDC, rcCell(lCell), 4, 15)
        Case 2: lRet = DrawEdge(lhDC, rcCell(lCell), 2, 15)
        Case 3
          hBrush = CreateSolidBrush(pvGetColor(lLineColor))
          lRet = SetRect(rcCell(lCell), rcCell(lCell).Left, rcCell(lCell).Top, rcCell(lCell).Right + pvEvaluate(lX < 6, 1, 0), rcCell(lCell).Bottom + pvEvaluate(lY < 6, 1, 0))
          lRet = FrameRect(lhDC, rcCell(lCell), hBrush)
          lRet = DeleteObject(hBrush)
      End Select
      lCell = lCell + 1
    Next lX
  Next lY
 'Draw week day cells
  For lCell = 0 To 6
    If (Not bShowWeekDays) Then Exit For
    hBrush = CreateSolidBrush(pvGetColor(lHeaderBack))
    lRet = SetRect(rcCell(lCell), (lBorderW + lOffset) + (lCell * lCellW), lBorderW + pvEvaluate(bShowMonth, 19, 0), (lBorderW + lOffset) + ((lCell + 1) * lCellW), lBorderW + lHeaderH)
    sBuffer = WeekdayName(lCell + 1, True, lFirstWeekDay)
    lRet = FillRect(lhDC, rcCell(lCell), hBrush)
    lRet = DeleteObject(hBrush)
    lRet = SetTextColor(lhDC, pvGetColor(lHeaderFore))
    lRet = DrawText(lhDC, sBuffer, Len(sBuffer), rcCell(lCell), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
    Select Case LineStyle
      Case 1: lRet = DrawEdge(lhDC, rcCell(lCell), 4, 15)
      Case 2: lRet = DrawEdge(lhDC, rcCell(lCell), 2, 15)
      Case 3
        hBrush = CreateSolidBrush(pvGetColor(lLineColor))
        lRet = SetRect(rcCell(lCell), rcCell(lCell).Left, rcCell(lCell).Top, rcCell(lCell).Right + pvEvaluate(lCell < 6, 1, 0), rcCell(lCell).Bottom + 1)
        lRet = FrameRect(lhDC, rcCell(lCell), hBrush)
        lRet = DeleteObject(hBrush)
    End Select
  Next lCell
 'Draw month name and year
  If (bShowMonth) Then
    hBrush = CreateSolidBrush(pvGetColor(lHeaderBack))
    lRet = SetRect(rcRect, lBorderW, lBorderW, ScaleWidth - lBorderW, lBorderW + 19)
    lRet = FillRect(lhDC, rcRect, hBrush)
    lRet = DeleteObject(hBrush)
    lRet = SetTextColor(lhDC, pvGetColor(lHeaderFore))
    sBuffer = MonthName(lMonth) & " " & lYear
    lHdrFont = SelectObject(lhDC, hHdrFont)
    lRet = DrawText(lhDC, sBuffer, Len(sBuffer), rcRect, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
    lRet = SelectObject(lhDC, lHdrFont)
    If (bMonthButtons) Then
      lRet = SetRect(rcButton(0), lBorderW + 2, lBorderW + 2, lBorderW + 17, lBorderW + 17)
      lRet = SetRect(rcButton(1), (ScaleWidth - lBorderW) - 17, lBorderW + 2, (ScaleWidth - lBorderW) - 2, lBorderW + 17)
      lSymFont = SelectObject(lhDC, hSymFont)
      lRet = DrawText(lhDC, "3", 1, rcButton(0), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      lRet = DrawText(lhDC, "4", 1, rcButton(1), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      lRet = SelectObject(lhDC, lSymFont)
    End If
    Select Case LineStyle
      Case 1: lRet = DrawEdge(lhDC, rcRect, 4, 15)
      Case 2: lRet = DrawEdge(lhDC, rcRect, 2, 15)
      Case 3
        lRet = SetRect(rcRect, lBorderW, lBorderW, ScaleWidth - lBorderW, lBorderW + 20)
        hBrush = CreateSolidBrush(pvGetColor(lLineColor))
        lRet = FrameRect(lhDC, rcRect, hBrush)
        lRet = DeleteObject(hBrush)
    End Select
  End If
 'Draw week numbers
  If (bShowWeekNum) Then
    For lCell = 48 To 54
      hBrush = CreateSolidBrush(pvGetColor(lHeaderBack))
      If (lCell < 54) Then
        lRet = ((lCell - 48) * lCellH) + (lBorderW + lHeaderH)
        lRet = SetRect(rcCell(lCell), lBorderW, lRet, lBorderW + pvEvaluate(lLineStyle = 3, 25, 24), lRet + lCellH - pvEvaluate((lCell = 53) And (lLineStyle = 3), 1, 0))
        lRet = DateSerial(lYear, lMonth, 1)
        lRet = DatePart("ww", lRet, lFirstWeekDay)
        sBuffer = lRet + (lCell - 48)
      Else
        lRet = SetRect(rcCell(lCell), lBorderW, lBorderW + pvEvaluate(bShowMonth, 19, 0), lBorderW + pvEvaluate(lLineStyle = 3, 25, 24), lBorderW + lHeaderH)
        sBuffer = ""
      End If
      lRet = FillRect(lhDC, rcCell(lCell), hBrush)
      lRet = DeleteObject(hBrush)
      lRet = SetTextColor(lhDC, pvGetColor(lOutDaysFore))
      lRet = DrawText(lhDC, sBuffer, Len(sBuffer), rcCell(lCell), DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
      Select Case LineStyle
        Case 1: lRet = DrawEdge(lhDC, rcCell(lCell), 4, 15)
        Case 2: lRet = DrawEdge(lhDC, rcCell(lCell), 2, 15)
        Case 3
          hBrush = CreateSolidBrush(pvGetColor(lLineColor))
          lRet = SetRect(rcCell(lCell), rcCell(lCell).Left, rcCell(lCell).Top, rcCell(lCell).Right + pvEvaluate(lCell < 6, 1, 0), rcCell(lCell).Bottom + 1)
          lRet = FrameRect(lhDC, rcCell(lCell), hBrush)
          lRet = DeleteObject(hBrush)
      End Select
    Next lCell
  End If
 'Draw border
  lRet = SetRect(rcRect, 0, 0, ScaleWidth, ScaleHeight)
  Select Case lBorderStyle
    Case 1: lRet = DrawEdge(lhDC, rcRect, 10, 15)
    Case 2: lRet = DrawEdge(lhDC, rcRect, 5, 15)
    Case 3: lRet = DrawEdge(lhDC, rcRect, 2, 15)
    Case 4: lRet = DrawEdge(lhDC, rcRect, 4, 15)
    Case 5: lRet = DrawEdge(lhDC, rcRect, 6, 15)
    Case 6
      hBrush = CreateSolidBrush(pvGetColor(lBorderColor))
      lRet = FrameRect(lhDC, rcRect, hBrush)
      lRet = DeleteObject(hBrush)
  End Select
 'Restore and delete fonts
  lRet = SelectObject(lhDC, lFont)
  lRet = DeleteObject(hFont)
  lRet = DeleteObject(hSymFont)
  lRet = DeleteObject(hHdrFont)
  lRet = DeleteObject(hSelFont)
 'Clear on-screen canvas, draw from the off-screen one and destroy'em
  Call Cls
  Call BitBlt(hDC, 0, 0, ScaleWidth, ScaleHeight, lhDC, 0, 0, vbSrcCopy)
  Call pvDestroy(lhDC, lhBmp, loBmp)
End Sub
'---pvIsThemeActive:
Private Function pvIsThemeActive() As Boolean
Dim hTheme As Long
Dim hModule As Long
Dim lRet As Long
  'If App.LogMode = 0 Then Exit Function
  hModule = LoadLibrary("UxTheme")
  If (hModule <> 0) Then
    lRet = GetProcAddress(hModule, "OpenThemeData")
    If (lRet <> 0) Then
      hTheme = OpenThemeData(hWnd, StrPtr("Window"))
      If (hTheme <> 0) Then pvIsThemeActive = True
      lRet = CloseThemeData(hTheme)
    End If
    lRet = FreeLibrary(hModule)
  End If
End Function
'---pvGetColor:
Private Function pvGetColor(ByVal Color As Long) As Long
  If (Color And &H80000000) Then
    pvGetColor = GetSysColor(Color And &H7FFFFFFF)
  Else
    pvGetColor = Color
  End If
End Function
'---pvBlendColor: Blends two colors and returns the resulting color
Private Function pvBlendColor(ByVal ColorFrom As Long, ByVal ColorTo As Long, Optional ByVal Alpha As Long = 128) As Long
Dim lCFrom As Long, lCTo As Long
Dim lSrcR As Long, lSrcG As Long, lSrcB As Long
Dim lDstR As Long, lDstG As Long, lDstB As Long
  lCFrom = pvGetColor(ColorFrom)
  lCTo = pvGetColor(ColorTo)
  lSrcR = lCFrom And &HFF
  lSrcG = (lCFrom And &HFF00&) \ &H100&
  lSrcB = (lCFrom And &HFF0000) \ &H10000
  lDstR = lCTo And &HFF
  lDstG = (lCTo And &HFF00&) \ &H100&
  lDstB = (lCTo And &HFF0000) \ &H10000
  pvBlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
End Function
'---pvGetDayByCell:
Private Function pvGetDayByCell(ByVal Cell As Long) As Long
Dim lWeekDay As Long
Dim lWeek As Long
Dim lFirstWeek As Long, lDays As Long
Dim lFirstYearDay As Long
Dim lFirstMonthDay As Long
Dim lLastMonthDay As Long
Dim sBuffer As String
  pvGetDayByCell = 0
  lFirstWeek = DateSerial(lYear, lMonth, 1)
  lFirstWeek = DatePart("ww", lFirstWeek, lFirstWeekDay)
  lFirstYearDay = DateSerial(lYear, 1, 1)
  lFirstYearDay = DatePart("w", lFirstYearDay, lFirstWeekDay)
  lFirstMonthDay = DateSerial(lYear, lMonth, 1)
  lFirstMonthDay = DatePart("w", lFirstMonthDay, lFirstWeekDay)
  lLastMonthDay = DateSerial(lYear, lMonth + 1, 1 - 1)
  lLastMonthDay = DatePart("d", lLastMonthDay, lFirstWeekDay)
  Call pvCellToLocation(Cell, lWeek, lWeekDay)
  If (lFirstMonthDay <> 1) Then lWeek = lWeek + 1
  lWeek = lFirstWeek + (lWeek - 1)
  lDays = ((lWeek - 2) * 7) + lWeekDay - lFirstYearDay
  sBuffer = DateSerial(lYear, 1, 1)
  sBuffer = DateAdd("d", lDays, sBuffer)
  sBuffer = DatePart("d", sBuffer, lFirstWeekDay)
  pvGetDayByCell = CLng(sBuffer)
End Function
'---pvGetCellByDay:
Private Function pvGetCellByDay(ByVal Day As Long) As Long
Dim lLastDay As Long
Dim lFirstCell As Long
Dim lLastCell As Long
  lLastDay = DateSerial(lYear, lMonth + 1, 1 - 1)
  lLastDay = DatePart("d", lLastDay, lFirstWeekDay)
  If (Day < 1) Or (lDay > lLastDay) Then Exit Function
  Call pvGetPeriodBounds(lFirstCell, lLastCell)
  pvGetCellByDay = lFirstCell + (Day - 1)
End Function
'---pvLocationToCell:
Private Sub pvLocationToCell(ByVal Row As Long, ByVal Column As Long, ByRef Cell As Long)
 'There's a very simple to do this :)
  Cell = (((Row - 1) * 7) + Column) + 6
End Sub
'---pvCellToLocation:
Private Sub pvCellToLocation(ByVal Cell As Long, ByRef Row As Long, ByRef Column As Long)
 'This way is faster and more accurate, and also doesn't requires any float operation
  Select Case Cell - 6
    Case Is <= 7:  Row = 1
    Case Is <= 14: Row = 2
    Case Is <= 21: Row = 3
    Case Is <= 28: Row = 4
    Case Is <= 35: Row = 5
    Case Is <= 42: Row = 6
  End Select
  Column = (Cell - 6) Mod 7
  If Column = 0 Then Column = 7
End Sub
'---pvGetActiveCell:
Private Function pvGetActiveCell(ByVal X As Long, ByVal Y As Long) As Long
Dim rcCell(48) As RECT
Dim lX As Long, lY As Long
Dim lCellW As Long, lCellH As Long
Dim lCell As Long
Dim lBorderW As Long
Dim lHeaderH As Long
Dim lOffset As Long
Dim lRet As Long
  pvGetActiveCell = 0
  Select Case lBorderStyle
    Case 0: lBorderW = 0
    Case 1, 2, 5: lBorderW = 2
    Case 3, 4, 6: lBorderW = 1
  End Select
  If (bShowMonth) Then lHeaderH = 19
  If (bShowWeekDays) Then lHeaderH = lHeaderH + 19
  If (bShowWeekNum) Then lOffset = 24
  lCellW = (ScaleWidth - ((lBorderW * 2) + lOffset)) \ 7
  lCellH = (ScaleHeight - ((lBorderW * 2) + lHeaderH)) \ 6
  lCell = 7
  For lY = 1 To 6
    For lX = 0 To 6
      lRet = SetRect(rcCell(lCell), (lBorderW + lOffset) + (lX * lCellW), (lBorderW + lHeaderH) + ((lY - 1) * lCellH), (lBorderW + lOffset) + ((lX + 1) * lCellW), (lBorderW + lHeaderH) + (lY * lCellH))
      If PtInRect(rcCell(lCell), X, Y) = 1 Then
        pvGetActiveCell = lCell
        Exit Function
      End If
      lCell = lCell + 1
    Next lX
  Next lY
End Function
'---pvIsWithinPeriod:
Private Function pvIsWithinPeriod(ByVal Cell As Long) As Boolean
Dim lFirstCell As Long
Dim lLastCell As Long
  Call pvGetPeriodBounds(lFirstCell, lLastCell)
  pvIsWithinPeriod = ((Cell >= lFirstCell) And (Cell <= lLastCell))
End Function
'---pvGetPeriodBounds:
Private Sub pvGetPeriodBounds(ByRef FirstCell As Long, ByRef LastCell As Long)
Static lFirstCell As Long
Static lLastCell As Long
Dim lLastDay As Long
Dim lCell As Long
Dim lDay As Long
  lDay = pvGetDayByCell(lFirstCell)
  If (lFirstCell > 0) And (lLastCell > 0) And (lDay = 1) Then GoTo OnlyReturn
  lFirstCell = 0
  lLastCell = 0
  lLastDay = DateSerial(lYear, lMonth + 1, 1 - 1)
  lLastDay = DatePart("d", lLastDay, lFirstWeekDay)
  For lCell = 7 To 48
    lDay = pvGetDayByCell(lCell)
    If (lFirstCell = 0) And (lDay = 1) Then
      lFirstCell = lCell
    End If
    If (lFirstCell > 0) And (lLastCell = 0) And (lDay = lLastDay) Then
      lLastCell = lCell
    End If
  Next lCell
OnlyReturn:
  FirstCell = lFirstCell
  LastCell = lLastCell
End Sub
'---pvCheckMonthMenu:
Private Sub pvCheckMonthMenu()
Dim lIndex As Long
  For lIndex = 1 To 12
    mnuMonth(lIndex).Checked = False
  Next lIndex
  mnuMonth(lMonth).Checked = True
End Sub
'---pvChangePeriod:
Private Sub pvChangePeriod()
Dim lPrevYear As Long
  lPrevYear = lYear
  If Not pvIsWithinPeriod(lActiveCell) Then
    If (lActiveCell < 15) Then
      lMonth = lMonth - 1
      If (lMonth < 1) Then
        lMonth = 12
        lYear = lYear - 1
      End If
    Else
      lMonth = lMonth + 1
      If (lMonth > 12) Then
        lMonth = 1
        lYear = lYear + 1
      End If
    End If
  End If
  If (lPrevYear <> lYear) Then Call pvUpdateYearsMenu
  lActiveCell = pvGetCellByDay(lDay)
End Sub
'---CreateFont:
Private Function CreateFont(Optional ByVal Name As String, Optional ByVal Size As Long = 8, Optional ByVal Charset As Long = 1, Optional ByVal Bold As Boolean, Optional ByVal Italic As Boolean, Optional ByVal Underline As Boolean, Optional ByVal Strikeout As Boolean, Optional ByVal Angle As Long) As Long
Dim lHeight As Long
Dim lWeight As Long
Dim sBuffer As String
  lHeight = -MulDiv(Size, GetDeviceCaps(hDC, 90), 72)
  If Bold Then lWeight = 700 Else lWeight = 400
  If Len(Name) > 32 Then Name = Left(Name, 32)
  sBuffer = Name & String(32 - Len(Name), vbNullChar)
  If Angle > 0 Then Angle = Angle * 10
  CreateFont = CreateFontA(lHeight, 0, Angle, Angle, lWeight, Italic, Underline, Strikeout, Charset, 0, 0, 0, 0, sBuffer)
End Function
'---pvCreate: Create bitmap and a DC to hold it
Private Sub pvCreate(ByRef hDC As Long, ByRef hBitmap As Long, ByRef hBitmapOld As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal bMono As Boolean = False)
Dim lhDC As Long
 'Create a DC to hold the bitmap
  If (bMono) Then
    hDC = CreateCompatibleDC(0)
    hBitmap = CreateCompatibleBitmap(hDC, Width, Height)
  Else
    lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    hDC = CreateCompatibleDC(lhDC)
    hBitmap = CreateCompatibleBitmap(lhDC, Width, Height)
  End If
 'Now create the bitmap
  If hBitmap = 0 Then GoTo CantCreateIt
 'Store the old bitmap and free resources
  hBitmapOld = SelectObject(hDC, hBitmap)
  Call DeleteDC(lhDC)
  Exit Sub
CantCreateIt:
 'Something went wrong, free resources
  Call DeleteDC(lhDC)
  Call pvDestroy(hDC, hBitmap, hBitmapOld)
End Sub
'---pvDestroy: Destroy a bitmap and the DC that holds it
Private Sub pvDestroy(ByRef hDC As Long, ByRef hBitmap As Long, ByRef hBitmapOld As Long)
 'Free resources
  If (hDC = 0) Then Exit Sub
  Call SelectObject(hDC, hBitmapOld)
  Call DeleteDC(hDC)
  Call DeleteObject(hBitmap)
 'Invalidate handles
  hBitmapOld = 0
  hBitmap = 0
  hDC = 0
End Sub
'---pvSetOwnerDrawn:
Private Sub pvSetOwnerDrawn(ByVal hMenu As Long)
Dim tMenuItemInfo As MENUITEMINFO
Dim lMenus As Long
Dim lMenu As Long
Dim lRet As Long
  lMenus = GetMenuItemCount(hMenu)
  For lMenu = 0 To lMenus - 1
    tMenuItemInfo.cbSize = Len(tMenuItemInfo)
    tMenuItemInfo.fMask = &H32
    tMenuItemInfo.cch = 127
    tMenuItemInfo.dwTypeData = String(128, vbNullChar)
    lRet = GetMenuItemInfo(hMenu, lMenu, 1, tMenuItemInfo)
    If Len(tMenuItemInfo.dwTypeData) > 1 Then sCaptions(tMenuItemInfo.wID) = Left(tMenuItemInfo.dwTypeData, InStr(1, tMenuItemInfo.dwTypeData, vbNullChar) - 1)
    tMenuItemInfo.fMask = &H30
    tMenuItemInfo.fType = tMenuItemInfo.fType Or &H100&
    tMenuItemInfo.dwItemData = lMenu + 1
    lRet = SetMenuItemInfo(hMenu, lMenu, 1, tMenuItemInfo)
  Next lMenu
End Sub
'---pvDrawMenuItem:
Private Sub pvDrawMenuItem(ByRef DIStruct As DRAWITEMSTRUCT)
Dim tNCM As NONCLIENTMETRICS
Dim rcMenu As RECT, rcText As RECT
Dim rcBar As RECT, bSelected As Boolean
Dim bChecked As Boolean, bDisabled As Boolean
Dim lhDC As Long, lhBmp As Long, loBmp As Long
Dim hFont As Long, lFont As Long
Dim hSymFont As Long, lSymFont As Long
Dim hBrush As Long, lRet As Long
Dim sBuffer As String, bXPStyle As Boolean
  Call pvCreate(lhDC, lhBmp, loBmp, DIStruct.rcItem.Right - DIStruct.rcItem.Left, DIStruct.rcItem.Bottom - DIStruct.rcItem.Top)
  bXPStyle = pvIsThemeActive
  bSelected = ((DIStruct.itemState And ODS_SELECTED) = ODS_SELECTED)
  bChecked = ((DIStruct.itemState And ODS_CHECKED) = ODS_CHECKED)
  bDisabled = ((DIStruct.itemState And ODS_GRAYED) = ODS_GRAYED)
  tNCM.cbSize = Len(tNCM)
  lRet = SystemParametersInfo(41, Len(tNCM), tNCM, 0)
  lRet = SetRect(rcMenu, 0, 0, DIStruct.rcItem.Right - DIStruct.rcItem.Left, DIStruct.rcItem.Bottom - DIStruct.rcItem.Top)
  lRet = SetBkMode(lhDC, 1)
  'Debug.Print DIStruct.itemID
  sBuffer = sCaptions(DIStruct.itemID) 'MonthName(DIStruct.itemData)
  If (bXPStyle) Then
    lRet = SetRect(rcBar, 0, 0, 24, rcMenu.Bottom)
  Else
    lRet = SetRect(rcBar, 0, 0, 18, rcMenu.Bottom)
  End If
  hFont = CreateFontIndirect(tNCM.lfMenuFont)
  hSymFont = CreateFont("Marlett", 10)
  lFont = SelectObject(lhDC, hFont)
  hBrush = CreateSolidBrush(pvGetColor(vbMenuBar))
  lRet = FillRect(lhDC, rcMenu, hBrush)
  lRet = DeleteObject(hBrush)
  If (bXPStyle) Then
    hBrush = CreateSolidBrush(pvGetColor(vb3DFace))
    lRet = FillRect(lhDC, rcBar, hBrush)
    lRet = DeleteObject(hBrush)
  End If
  If (bSelected) Then
    If (bXPStyle) Then
      hBrush = CreateSolidBrush(pvBlendColor(vbHighlight, vbWindowBackground, 70))
      lRet = FillRect(lhDC, rcMenu, hBrush)
      lRet = DeleteObject(hBrush)
      hBrush = CreateSolidBrush(pvGetColor(vbHighlight))
      lRet = FrameRect(lhDC, rcMenu, hBrush)
      lRet = DeleteObject(hBrush)
      lRet = SetTextColor(lhDC, pvGetColor(vbMenuText))
    Else
      hBrush = CreateSolidBrush(pvGetColor(vbHighlight))
      lRet = FillRect(lhDC, rcMenu, hBrush)
      lRet = DeleteObject(hBrush)
      lRet = SetTextColor(lhDC, pvGetColor(vbHighlightText))
    End If
  Else
    lRet = SetTextColor(lhDC, pvGetColor(vbMenuText))
  End If
  If (bChecked) Then
    If (bXPStyle) Then
      lRet = SetRect(rcBar, 1, 1, 22, rcMenu.Bottom - 1)
      hBrush = CreateSolidBrush(pvBlendColor(vbHighlight, vbWindowBackground, 30))
      lRet = FillRect(lhDC, rcBar, hBrush)
      lRet = DeleteObject(hBrush)
      hBrush = CreateSolidBrush(pvGetColor(vbHighlight))
      lRet = FrameRect(lhDC, rcBar, hBrush)
      lRet = DeleteObject(hBrush)
      lRet = SetTextColor(lhDC, pvGetColor(vbMenuText))
      lRet = SetRect(rcBar, 0, 0, 24, rcMenu.Bottom)
      lSymFont = SelectObject(lhDC, hSymFont)
      lRet = DrawText(lhDC, "a", 1, rcBar, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      lRet = SelectObject(lhDC, lSymFont)
    Else
      lSymFont = SelectObject(lhDC, hSymFont)
      lRet = DrawText(lhDC, "a", 1, rcBar, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
      lRet = SelectObject(lhDC, lSymFont)
    End If
  End If
  If (bXPStyle) Then
    lRet = SetRect(rcText, rcBar.Right + 6, 0, rcMenu.Right, rcMenu.Bottom)
  Else
    lRet = SetRect(rcText, rcBar.Right + 2, 0, rcMenu.Right, rcMenu.Bottom)
  End If
  lRet = DrawText(lhDC, sBuffer, Len(sBuffer), rcText, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS)
  lRet = SelectObject(lhDC, lFont)
  lRet = DeleteObject(hSymFont)
  lRet = DeleteObject(hFont)
  Call BitBlt(DIStruct.hDC, DIStruct.rcItem.Left, DIStruct.rcItem.Top, DIStruct.rcItem.Right - DIStruct.rcItem.Left, DIStruct.rcItem.Bottom - DIStruct.rcItem.Top, lhDC, 0, 0, vbSrcCopy)
  Call pvDestroy(lhDC, lhBmp, loBmp)
End Sub
'---pvEvaluate:
Private Function pvEvaluate(Expression, IfTrue, IfFalse) As Variant
  If (Expression) Then pvEvaluate = IfTrue Else pvEvaluate = IfFalse
End Function
'---pvUpdateYearsMenu:
Private Sub pvUpdateYearsMenu()
Dim lIndex As Long
  For lIndex = 1 To 11
    If (lIndex < 6) Then
      mnuYear(lIndex).Caption = lYear - (6 - lIndex)
    ElseIf (lIndex = 6) Then
      mnuYear(lIndex).Caption = lYear
    ElseIf (lIndex > 6) Then
      mnuYear(lIndex).Caption = lYear + (lIndex - 6)
    End If
  Next lIndex
End Sub
'***************************************************************************************
' Menu Events
'***************************************************************************************
'---mnuMonth_Click:
Private Sub mnuMonth_Click(Index As Integer)
Dim lLastDay As Long
  lMonth = Index
  Call pvCheckMonthMenu
  lLastDay = DateSerial(lYear, lMonth + 1, 1 - 1)
  lLastDay = DatePart("d", lLastDay, lFirstWeekDay)
  If (lDay > lLastDay) Then lDay = lLastDay
  lActiveCell = pvGetCellByDay(lDay)
  Call pvDrawControl
  sActiveDate = DateSerial(lYear, lMonth, lDay)
  RaiseEvent Change
End Sub
'---mnuYear_Click:
Private Sub mnuYear_Click(Index As Integer)
Dim lNewYear As Long
Dim lLastDay As Long
  If (Index < 6) Then
    lNewYear = lYear - (6 - Index)
  ElseIf (Index = 6) Then
    lNewYear = lYear
  ElseIf (Index > 6) Then
    lNewYear = lYear + (Index - 6)
  End If
  lYear = lNewYear
  lLastDay = DateSerial(lYear, lMonth + 1, 1 - 1)
  lLastDay = DatePart("d", lLastDay, lFirstWeekDay)
  If (lDay > lLastDay) Then lDay = lLastDay
  lActiveCell = pvGetCellByDay(lDay)
  Call pvUpdateYearsMenu
  sActiveDate = DateSerial(lYear, lMonth, lDay)
  Call pvDrawControl
  RaiseEvent Change
End Sub
'***************************************************************************************
' UserControl Events
'***************************************************************************************
'---UserControl_Initialize:
Private Sub UserControl_Initialize()
Dim lIndex As Long
  bIsLeapYear = (DatePart("yyyy", Date) Mod 4 = 0)
  For lIndex = 1 To 12
    mnuMonth(lIndex).Caption = MonthName(lIndex)
  Next lIndex
End Sub
'---UserControl_Terminate:
Private Sub UserControl_Terminate()
  If (bRunning) Then
    Call sc_Terminate
  End If
  bRunning = False
  bPropLoaded = False
End Sub
'---UserControl_EnterFocus:
Private Sub UserControl_EnterFocus()
  bHasFocus = True
  If (Not bMouseDown) Then Call pvDrawControl
End Sub
'---UserControl_ExitFocus:
Private Sub UserControl_ExitFocus()
  bHasFocus = False
  bMouseDown = False
  Call pvDrawControl
End Sub
'---UserControl_KeyDown:
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lRow As Long, lColumn As Long
Dim lRet As Long
  Call pvCellToLocation(lActiveCell, lRow, lColumn)
  Select Case KeyCode
    Case vbKeyUp:    lRow = lRow - 1
    Case vbKeyDown:  lRow = lRow + 1
    Case vbKeyLeft:  lColumn = lColumn - 1
    Case vbKeyRight: lColumn = lColumn + 1
    Case Else:       Exit Sub
  End Select
  Call pvLocationToCell(lRow, lColumn, lRet)
  If (lRet < 7) Or (lRet > 48) Then Exit Sub
  lActiveCell = lRet
  lDay = pvGetDayByCell(lRet)
  Call pvChangePeriod
  Call pvCheckMonthMenu
  Call pvDrawControl
  sActiveDate = DateSerial(lYear, lMonth, lDay)
  RaiseEvent Change
End Sub
'---UserControl_MouseDown:
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bMouseDown = True
End Sub
'---UserControl_MouseUp:
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rcMonth As RECT
Dim rcYear As RECT
Dim lTextW As Long
Dim lTextH As Long
Dim lWidth As Long
Dim rcButton(1) As RECT
Dim lNewCell As Long
Dim lBorderW As Long
Dim lRet As Long
  Select Case lBorderStyle
    Case 0: lBorderW = 0
    Case 1, 2, 5: lBorderW = 2
    Case 3, 4, 6: lBorderW = 1
  End Select
  Font.Bold = True
  lTextW = TextWidth(MonthName(lMonth) & " " & lYear)
  lTextH = TextHeight(MonthName(lMonth) & " " & lYear)
  lWidth = ScaleWidth - (lBorderW * 2)
  lRet = SetRect(rcMonth, (lWidth \ 2) - (lTextW \ 2), lBorderW, ((lWidth \ 2) - (lTextW \ 2)) + TextWidth(MonthName(lMonth)), lBorderW + 19)
  lRet = SetRect(rcYear, ((lWidth \ 2) + (lTextW \ 2)) - TextWidth(lYear), lBorderW, (lWidth \ 2) + (lTextW \ 2), lBorderW + 19)
  Font.Bold = False
  If (PtInRect(rcMonth, X, Y) = 1) And (Button <> vbMiddleButton) Then
    Call PopupMenu(mnuMonths, , rcMonth.Left, lBorderW + 18)
  ElseIf (PtInRect(rcYear, X, Y) = 1) And (Button <> vbMiddleButton) Then
    Call PopupMenu(mnuYears, , rcYear.Left, lBorderW + 18)
  End If
  If (Button = vbLeftButton) Then
    lRet = SetRect(rcButton(0), lBorderW + 2, lBorderW + 2, lBorderW + 17, lBorderW + 17)
    lRet = SetRect(rcButton(1), (ScaleWidth - lBorderW) - 17, lBorderW + 2, (ScaleWidth - lBorderW) - 2, lBorderW + 17)
    If (PtInRect(rcButton(0), X, Y) = 1) And (bMonthButtons) And (bShowMonth) Then
    lMonth = lMonth - 1
      If (lMonth < 1) Then
        lMonth = 12
        lYear = lYear - 1
      End If
      Call mnuMonth_Click(CInt(lMonth))
      Exit Sub
    End If
    If (PtInRect(rcButton(1), X, Y) = 1) And (bMonthButtons) And (bShowMonth) Then
      lMonth = lMonth + 1
      If (lMonth > 12) Then
        lMonth = 1
        lYear = lYear + 1
      End If
      Call mnuMonth_Click(CInt(lMonth))
      Exit Sub
    End If
    lNewCell = pvGetActiveCell(X, Y)
    If (lNewCell = 0) Then Exit Sub
    lActiveCell = lNewCell
    lDay = pvGetDayByCell(lActiveCell)
    Call pvChangePeriod
    Call pvCheckMonthMenu
    Call pvDrawControl
    sActiveDate = DateSerial(lYear, lMonth, lDay)
    RaiseEvent Change
    RaiseEvent Click
  End If
  Exit Sub
  bMouseDown = False
End Sub
'---UserControl_InitProperties:
Private Sub UserControl_InitProperties()
  sActiveDate = Date
  lLineStyle = 1
  lBorderStyle = 1
  lFocusStyle = 1
  lLineColor = vb3DShadow
  lBorderColor = vb3DShadow
  lFirstWeekDay = vbSunday
  lHeaderBack = vb3DShadow
  lHeaderFore = vb3DHighlight
  lDaysBack = vbWindowBackground
  lDaysFore = vbWindowText
  lOutDaysBack = vbButtonFace
  lOutDaysFore = vbGrayText
  lSelectBack = vb3DLight
  lSelectFore = vbButtonText
  lActiveDayFore = vbHighlight
  lYear = DatePart("yyyy", sActiveDate)
  lMonth = DatePart("m", sActiveDate)
  lDay = DatePart("d", sActiveDate)
  bShowMonth = True
  bShowWeekDays = True
  bMonthButtons = True
  bShowWeekNum = False
  Set UserControl.Font = Ambient.Font
  Set tSelFont = Ambient.Font
  Set tHdrFont = Ambient.Font
  tSelFont.Bold = True
  tHdrFont.Bold = True
  bPropLoaded = True
End Sub
'---UserControl_ReadProperties:
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    Set tSelFont = .ReadProperty("ActiveDayFont", Ambient.Font)
    Set tHdrFont = .ReadProperty("MonthNameFont", Ambient.Font)
    tSelFont.Bold = .ReadProperty("SelFontBold", True)
    tHdrFont.Bold = .ReadProperty("HdrFontBold", True)
    sActiveDate = .ReadProperty("ActiveDate", Date)
    lLineStyle = .ReadProperty("LineStyle", 1)
    lBorderStyle = .ReadProperty("BorderStyle", 1)
    lFocusStyle = .ReadProperty("FocusStyle", 1)
    lFirstWeekDay = .ReadProperty("FirstDayOfWeek", vbUseSystem)
    lLineColor = .ReadProperty("LineColor", vb3DShadow)
    lBorderColor = .ReadProperty("BorderColor", vb3DShadow)
    lHeaderBack = .ReadProperty("HeaderBackColor", vb3DShadow)
    lHeaderFore = .ReadProperty("HeaderForeColor", vb3DHighlight)
    lDaysBack = .ReadProperty("DaysBackColor", vbWindowBackground)
    lDaysFore = .ReadProperty("DaysForeColor", vbWindowText)
    lOutDaysBack = .ReadProperty("OutDaysBackColor", vbButtonFace)
    lOutDaysFore = .ReadProperty("OutDaysForeColor", vbGrayText)
    lActiveDayFore = .ReadProperty("ActiveDayForeColor", vbHighlight)
    bShowMonth = .ReadProperty("ShowMonth", True)
    bShowWeekDays = .ReadProperty("ShowWeekDays", True)
    bMonthButtons = .ReadProperty("MonthChangeButtons", True)
    bShowWeekNum = .ReadProperty("ShowWeekNumber", False)
  End With
  lYear = DatePart("yyyy", sActiveDate)
  lMonth = DatePart("m", sActiveDate)
  lDay = DatePart("d", sActiveDate)
  lActiveCell = pvGetCellByDay(lDay)
  mnuMonth(lMonth).Checked = True
  bPropLoaded = True
  bRunning = Ambient.UserMode
  If (bRunning) Then Set UserControl.Picture = Nothing
  Call pvUpdateYearsMenu
  Call pvDrawControl
  If (bRunning) And (CTL_IDE_SUBCLASS) Then
    Call sc_Subclass(hWnd)
    Call sc_AddMsg(hWnd, WM_THEMECHANGED)
    Call sc_AddMsg(hWnd, WM_SYSCOLORCHANGE)
    Call sc_AddMsg(hWnd, WM_INITMENUPOPUP)
    Call sc_AddMsg(hWnd, WM_DRAWITEM)
    Call sc_AddMsg(hWnd, WM_MEASUREITEM)
  End If
End Sub
'---UserControl_WriteProperties:
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call .WriteProperty("ActiveDayFont", tSelFont, Ambient.Font)
    Call .WriteProperty("MonthNameFont", tHdrFont, Ambient.Font)
    Call .WriteProperty("SelFontBold", tSelFont.Bold, True)
    Call .WriteProperty("HdrFontBold", tHdrFont.Bold, True)
    Call .WriteProperty("ActiveDate", sActiveDate, Date)
    Call .WriteProperty("LineStyle", lLineStyle, 1)
    Call .WriteProperty("BorderStyle", lBorderStyle, 1)
    Call .WriteProperty("FocusStyle", lFocusStyle, 1)
    Call .WriteProperty("FirstDayOfWeek", lFirstWeekDay, vbUseSystem)
    Call .WriteProperty("LineColor", lLineColor, vb3DShadow)
    Call .WriteProperty("BorderColor", lBorderColor, vb3DShadow)
    Call .WriteProperty("HeaderBackColor", lHeaderBack, vb3DShadow)
    Call .WriteProperty("HeaderForeColor", lHeaderFore, vb3DHighlight)
    Call .WriteProperty("DaysBackColor", lDaysBack, vbWindowBackground)
    Call .WriteProperty("DaysForeColor", lDaysFore, vbWindowText)
    Call .WriteProperty("OutDaysBackColor", lOutDaysBack, vbButtonFace)
    Call .WriteProperty("OutDaysForeColor", lOutDaysFore, vbGrayText)
    Call .WriteProperty("ActiveDayForeColor", lActiveDayFore, vbHighlight)
    Call .WriteProperty("ShowMonth", bShowMonth, True)
    Call .WriteProperty("ShowWeekDays", bShowWeekDays, True)
    Call .WriteProperty("MonthChangeButtons", bMonthButtons, True)
    Call .WriteProperty("ShowWeekNumber", bShowWeekNum, False)
  End With
End Sub
'---UserControl_Resize:
Private Sub UserControl_Resize()
Static lSobX As Single
Static lSobY As Single
Dim lBorderW As Long
Dim lHeaderH As Long
Dim lOffset As Long
  If (bIgnoreEvent) Then Exit Sub
  Select Case lBorderStyle
    Case 0: lBorderW = 0
    Case 1, 2, 5: lBorderW = 2
    Case 3, 4, 6: lBorderW = 1
  End Select
  bIgnoreEvent = True
  If (lSobX > 0) Then Width = ScaleX(ScaleWidth + lSobX, vbPixels, vbTwips)
  If (lSobY > 0) Then Height = ScaleY(ScaleHeight + lSobY, vbPixels, vbTwips)
  If (bShowMonth) Then lHeaderH = 19
  If (bShowWeekDays) Then lHeaderH = lHeaderH + 19
  If (bShowWeekNum) Then lOffset = 24
  lSobX = (ScaleWidth - ((lBorderW * 2) + lOffset)) Mod 7
  lSobY = (ScaleHeight - ((lBorderW * 2) + lHeaderH)) Mod 6
  If (lSobX > 0) Then Width = ScaleX(ScaleWidth - lSobX, vbPixels, vbTwips)
  If (lSobY > 0) Then Height = ScaleY(ScaleHeight - lSobY, vbPixels, vbTwips)
  bIgnoreEvent = False
  Call pvDrawControl
End Sub
'***************************************************************************************
' SelfSub Code
'***************************************************************************************
Private Function sc_Subclass(ByVal lng_hWnd As Long, Optional ByVal lParamUser As Long = 0, Optional ByVal nOrdinal As Long = 1, Optional ByVal oCallback As Object = Nothing, Optional ByVal bIdeSafety As Boolean = True) As Boolean
Const CODE_LEN      As Long = 260
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))
Const PAGE_RWX      As Long = &H40&
Const MEM_COMMIT    As Long = &H1000&
Const MEM_RELEASE   As Long = &H8000&
Const IDX_EBMODE    As Long = 3
Const IDX_CWP       As Long = 4
Const IDX_SWL       As Long = 5
Const IDX_FREE      As Long = 6
Const IDX_BADPTR    As Long = 7
Const IDX_OWNER     As Long = 8
Const IDX_CALLBACK  As Long = 10
Const IDX_EBX       As Long = 16
Const SUB_NAME      As String = "sc_Subclass"
Dim nAddr         As Long
Dim nID           As Long
Dim nMyID         As Long
  If IsWindow(lng_hWnd) = 0 Then
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If
  nMyID = GetCurrentProcessId
  GetWindowThreadProcessId lng_hWnd, nID
  If nID <> nMyID Then
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  If oCallback Is Nothing Then
    Set oCallback = Me
  End If
  nAddr = zAddressOf(oCallback, nOrdinal)
  If nAddr = 0 Then
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
  If z_Funk Is Nothing Then
    Set z_Funk = New Collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&
    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")
  End If
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)
  If z_ScMem <> 0 Then
    On Error GoTo CatchDoubleSub
      z_Funk.Add z_ScMem, "h" & lng_hWnd
    On Error GoTo 0
    If bIdeSafety Then
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")
    End If
    z_Sc(IDX_EBX) = z_ScMem
    z_Sc(IDX_HWND) = lng_hWnd
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)
    z_Sc(IDX_CALLBACK) = nAddr
    z_Sc(IDX_PARM_USER) = lParamUser
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)
    If nAddr = 0 Then
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
    z_Sc(IDX_WNDPROC) = nAddr
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN
    sc_Subclass = True
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  Exit Function
CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE
End Function
'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i As Long
  If Not (z_Funk Is Nothing) Then
    With z_Funk
      For i = .Count To 1 Step -1
        z_ScMem = .Item(i)
        If IsBadCodePtr(z_ScMem) = 0 Then
          sc_UnSubclass zData(IDX_HWND)
        End If
      Next i
    End With
    Set z_Funk = Nothing
  End If
End Sub
'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
      zData(IDX_SHUTDOWN) = -1
      zDelMsg ALL_MESSAGES, IDX_BTABLE
      zDelMsg ALL_MESSAGES, IDX_ATABLE
    End If
    z_Funk.Remove "h" & lng_hWnd
  End If
End Sub
'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
    If When And MSG_BEFORE Then
      zAddMsg uMsg, IDX_BTABLE
    End If
    If When And MSG_AFTER Then
      zAddMsg uMsg, IDX_ATABLE
    End If
  End If
End Sub
'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
    If When And MSG_BEFORE Then
      zDelMsg uMsg, IDX_BTABLE
    End If
    If When And MSG_AFTER Then
      zDelMsg uMsg, IDX_ATABLE
    End If
  End If
End Sub
'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
    sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam)
  End If
End Function
'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
    sc_lParamUser = zData(IDX_PARM_USER)
  End If
End Property
'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
    zData(IDX_PARM_USER) = NewValue
  End If
End Property
'-The following routines are exclusively for the sc_ subclass routines----------------------------
'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long
  Dim nBase  As Long
  Dim i      As Long
  nBase = z_ScMem
  z_ScMem = zData(nTable)
  If uMsg = ALL_MESSAGES Then
    nCount = ALL_MESSAGES
  Else
    nCount = zData(0)
    If nCount >= MSG_ENTRIES Then
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If
    For i = 1 To nCount
      If zData(i) = 0 Then
        zData(i) = uMsg
        GoTo Bail
      ElseIf zData(i) = uMsg Then
        GoTo Bail
      End If
    Next i
    nCount = i
    zData(nCount) = uMsg
  End If
  zData(0) = nCount
Bail:
  z_ScMem = nBase
End Sub
'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long
  Dim nBase  As Long
  Dim i      As Long
  nBase = z_ScMem
  z_ScMem = zData(nTable)
  If uMsg = ALL_MESSAGES Then
    zData(0) = 0
  Else
    nCount = zData(0)
    For i = 1 To nCount
      If zData(i) = uMsg Then
        zData(i) = 0
        GoTo Bail
      End If
    Next i
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
Bail:
  z_ScMem = nBase
End Sub
'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub
'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zFnAddr
End Function
'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch
    z_ScMem = z_Funk("h" & lng_hWnd)
    zMap_hWnd = z_ScMem
  End If
  Exit Function
Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function
'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte
  Dim bVal  As Byte
  Dim nAddr As Long
  Dim i     As Long
  Dim j     As Long
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4
  If Not zProbe(nAddr + &H1C, i, bSub) Then
    If Not zProbe(nAddr + &H6F8, i, bSub) Then
      If Not zProbe(nAddr + &H7A4, i, bSub) Then
        Exit Function
      End If
    End If
  End If
  i = i + 4
  j = i + 1024
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4
    If IsBadCodePtr(nAddr) Then
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4
      Exit Do
    End If
    RtlMoveMemory VarPtr(bVal), nAddr, 1
    If bVal <> bSub Then
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4
      Exit Do
    End If
    i = i + 4
  Loop
End Function
'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  nAddr = nStart
  nLimit = nAddr + 32
  Do While nAddr < nLimit
    RtlMoveMemory VarPtr(nEntry), nAddr, 4
    If nEntry <> 0 Then
      RtlMoveMemory VarPtr(bVal), nEntry, 1
      If bVal = &H33 Or bVal = &HE9 Then
        nMethod = nAddr
        bSub = bVal
        zProbe = True
        Exit Function
      End If
    End If
    nAddr = nAddr + 4
  Loop
End Function
Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property
Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property
'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
Dim lValue As Long, lNewValue As Long
Dim tMeISt As MEASUREITEMSTRUCT
Dim tDrISt As DRAWITEMSTRUCT
Dim tNCM As NONCLIENTMETRICS
Dim hFont As Long, lFont As Long
Dim lRet As Long, sBuffer As String
  Select Case uMsg
    Case WM_SYSCOLORCHANGE, WM_THEMECHANGED
      Call pvDrawControl
    Case WM_INITMENUPOPUP
      Call pvSetOwnerDrawn(wParam)
    Case WM_MEASUREITEM
      Call CopyMemory(tMeISt, ByVal lParam, Len(tMeISt))
      tNCM.cbSize = Len(tNCM)
      lRet = SystemParametersInfo(41, Len(tNCM), tNCM, 0)
      hFont = CreateFontIndirect(tNCM.lfMenuFont)
      lFont = SelectObject(hDC, hFont)
      sBuffer = sCaptions(tMeISt.itemID)
      If (pvIsThemeActive) Then
        tMeISt.itemWidth = tMeISt.itemWidth + 34 + TextWidth(sBuffer)
        tMeISt.itemHeight = tNCM.iMenuHeight + 3
      Else
        tMeISt.itemWidth = tMeISt.itemWidth + 26 + TextWidth(sBuffer)
        tMeISt.itemHeight = tNCM.iMenuHeight
      End If
      lRet = SelectObject(hDC, lFont)
      lRet = DeleteObject(hFont)
      Call CopyMemory(ByVal lParam, tMeISt, Len(tMeISt))
    Case WM_DRAWITEM
      Call CopyMemory(tDrISt, ByVal lParam, Len(tDrISt))
      Call pvDrawMenuItem(tDrISt)
  End Select
End Sub
