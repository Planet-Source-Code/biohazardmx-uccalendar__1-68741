 ~-~-~-~-~-~-~-~-~-~
  ucCalendar Readme
 -~-~-~-~-~-~-~-~-~-

  Name:      ucCalendar
  Version:   0.9 RC1 (Release Candidate 1)
  Date:      25/05/2007
  Author:    BioHazardMX
  Details:   Calendar Control (self-contained, full API drawn)
  Requires:  Win32 Libraries (User32, GDI32, Kernel32)

 --------------
  Introduction
 --------------
 
  ucCalendar is an small and self contained user control that provides basic calendar
  functionality. With support for mouse and keyboard control, multiple border, line 
  and even selection styles you can integrate it with ANY user interface.
  
  ucCalendar will be a replacement for the MontView control, it has almost all its 
  features and we're very close to a final version. The current release is only a 
  small preview, an almost-finished work-in-progress, but some of the final functionality
  is not already in. Also there are a few bugs that must be fixed and a lot of 
  testing to be done.
 
  If you're interested in contributing with the development, please contact me and
  send me any changes that you made. Also your feedback is welcome, and bug reports
  are indispensable. For a list of the pending features, refer to the 'ToDo' section.

  I am releasing this public Beta because I need to ensure its proper operation under
  diferent ambients (I have tested it under WinXP and a VirtualPC WinME but it's still
  the same machine). If you have an old system (Win95, 98) or a newer one (Vista) please
  try to run the sample project and send me your feedback.

  Thanks for trying ucCalendar!

 ---------
  Changes
 ---------

 Version 0.9 RC 1 (1,513 lines)
  * New toolbox icon
  * Release Candidate 1, the next version may be a final release
  * Added descriptions for procedures, properties and events (for Properties window)
  * Added pvGetCellByDay and pvGetDayByCell procedures (simplifying other procedures)
  * Fixed small issue with months that start in Sunday (or the specified first week day)
  * Fixed automatic month changer, the last day will be clipped accordingly to each month
  * Fixed memory leak in pvDrawControl (not deleting the symbol font due to a mistype)
  * Updated pvGetPeriodBounds, Itll only search the period bounds when month has changed
  * Added SelfSub-powered subclassing to detect theme and color changes to redraw control
  * Updated month change menu, now it uses Office XP style under WinXP when themes are activated
  * Fixed: ShadedRect + SimpleLine = Overlapped borders in the 7th column's right border
  * Added 'ShowWeekNumber' property, toggle week numbers on the left of the calendar
  * Fixed resizing behaviour, finally I've found a working (at least for me) workaround
  * Added a new adaptive year change menu, it shows the previous/next 5 years to choose from!
  * Fixed menus appearing in MouseDown event (now they appear with MouseUp)
  * Fixed stupid bug with month change buttons even when the month name bar where not visible!

 Version 0.8 Alpha (602 lines)
  * First alpha version, almost feature-complete
  * Added autosizing feature, now the control is adjusted to show only complete cells
  * Fixed issue when changing border style, now the property fires the Resize event
  * Added support for non-sunday-starting weeks (through FirstDayOfWeek property)
  * Added ShowMonth and ShowDayNames properties
  * Fixed small issue in pvLocationToCell parameters (they where inverted)
  * Updated drawing procedures, now using an off-screen DC for faster rendering
  * Added a basic keyboard interface
  * Added automatic month change when clicking previous/next month days
  * Added month change menu (click menu name)
  * Added month change buttons (optional, see MonthChangeButtons property )
  * Added FocusStyle property to customize the selection rectangle style
  * Fixed small glitch when drawing after receiving focus with mouse down
  * Code rearranged, unused stuff was removed and procedures have been simplified

 ------
  ToDo
 ------

  * Maybe year change buttons (with the adaptive menu this isn't required)
  * Maybe visual themes (Win9x, WinXP, OfficeXP, Office 2003, MacOSX, etc)
  * Highligthing of specific dates (not only the active) to show several events in one
    calendar (like Rainlendar does)
  * More MonthView functionality (week numbers[DONE], multiple selections, etc)

  Can you help me? Have you already added one of these features?
    If so, please contact me to update the control.

 ---------
  License
 ---------

  You may use this control and/or any part of its source code freely as long as you
  agree that no warranty or liability of any kind is expressed or implied. You may not
  remove any copyright notices and also you may not claim ownership unless you have
  substantially altered the functionality or features.

  This software uses the SelfSub subclassing system from Paul Caton.
  Thanks to Paul Caton for this excellent code, find out more at:
   http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1

  Copyright © 2007 BioHazardMX.
  Proudly made in Mexico.