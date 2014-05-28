OPT (Outlook Prayer Times)
==========================

Ever wanted to reserve Islamic prayer times in your Outlook calendar? Ever wanted to get reminded about prayer times like normal Outlook meetings? OPT is the right tool for you. It calculates Islamic prayer times and book a corresponding meeting request in your Outlook calendar.

OPT is developed as Outlook macros using open source VBA. It is based on prayer time code posted on [PrayerTimes.org] [1].

Version
-------

0.9

Installation
------------

1. Download the code directly from GitHub and save it to a local directory.
2. Enable the developer tab as highlighted [here] [2].
3. Enable Outlook to run macros:

 1. Go to developer tab.
 2. Click on Macros Security.
 3. In Macros Settings, select Notifications for all macros.
 4. Click Ok.

4. Still in developer tab, import OPT code as following:

 1. Click Visual Basic.
 2. Right-click Microsoft Outlook Objects, Import File..., then select
    ThisOutlookSession.bas from local directory. Click open.
 3. Right-click Modules, Import File..., then select FillCalendar.bas from
    local directory. Click open.
 4. Right-click Class Modules, Import File..., then select ItemEvents.bas from
    local directory. Click open.
 5. Repeat step 4, for PrayerTime.bas.
 6. Double click on FillCalendar to open it. Adjust the code shown blow to reflect
    your country and time zone. Please note that DayLightSaving is not currently
    used.
  ```vba
  ' Edit these to reflect your country
  Public Const TimeZone As Double = 3#
  Public Const DayLightSaving As Double = 0
  Public Const Latitude As Double = 30.0566
  Public Const Longitude As Double = 31.2262
  ```
 7. Do not forget to save by clicking the save icon.
 8. Close Microsoft Visual Basic for Applications window.
 9. Restart Outlook.
 10. At startup, Outlook will ask to enable macros. Grant it.
 11. Switch to Calendar view, to verify the new green prayer times meetings.


License
-------

LGPL

[1]:http://praytimes.org/wiki/Code
[2]:http://msdn.microsoft.com/en-us/library/bb608625.aspx
