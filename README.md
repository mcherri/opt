OPT (Outlook Prayer Times)
==========================

Ever wanted to reserve Islamic prayer times in your Outlook calendar? Ever wanted
to get reminded about prayer times like normal Outlook meetings? OPT is the right
tool for you. It calculates Islamic prayer times and book a corresponding meeting
request in your Outlook calendar.

OPT is developed as Outlook macros using open source VBA. It is based on prayer
time code posted on [PrayerTimes.org] [1].

Version
-------

0.9

Installation
------------

### Easy but Risky Installation Steps

1. Please note that this method is risky as highlighted in this [KB] [2]. This
   method works if you do not have existing macros in Outlook.
2. Enable the developer tab as highlighted [here] [3].
3. Enable Outlook to run macros:

 1. Go to developer tab.
 2. Click on Macros Security.
 3. In Macros Settings, select Notifications for all macros.
 4. Click Ok.

4. Close Outlook if it is open.
5. Open %appdata%\Microsoft\Outlook
6. Take a backup of VbaProject.OTM if available.
7. Download VbaProject.OTM from OPT and replace your default VbaProject.OTM.
8. Restart Outlook.
9. At startup, Outlook will ask to enable macros. Grant it.
10. Switch to Calendar view, to verify the new green prayer times meetings'.
11. If the prayer times do not reflect your current location look at the long
    installation steps.

### Long Installation Steps

1. Download the code directly from GitHub and save it to a local directory.
2. Enable the developer tab as highlighted [here] [3].
3. Enable Outlook to run macros:

 1. Go to developer tab.
 2. Click on Macros Security.
 3. In Macros Settings, select Notifications for all macros.
 4. Click Ok.

4. Still in developer tab, import OPT code as following:

 1. Click Visual Basic.
 2. Expand Microsoft Outlook Objects, double click ThisOutlookSession, and
    copy and paste ThisOutlookSession.bas into the newly opened code window.
 3. From the menu, click Insert -> Module. In the module properties, change
    the module name to FillCalendar. Copy and paste FillCalendar.bas into the
    newly opened code window.
 4. From the menu click Insert -> Class Module. In the class module properties,
    change the class module name to ItemEvents. Copy and paste ItemEvents.bas
    into the newly opened code window.
 5. Repeat step iv, for PrayerTime.bas.
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
 11. Switch to Calendar view, to verify the new green prayer times meetings'.

### Re-enable Full Macros Security

1. Perform the steps in this [article] [4] to self-sign you macros.
2. If you face any issues try to follow the steps in the most helpful reply
   in this [thread] [5].

License
-------

LGPL

[1]:http://praytimes.org/wiki/Code
[2]:http://support.microsoft.com/KB/229911
[3]:http://msdn.microsoft.com/en-us/library/bb608625.aspx
[4]:http://www.howto-outlook.com/howto/selfcert.htm
[5]:http://answers.microsoft.com/en-us/office/forum/office_2010-customize/how-can-i-digitally-sign-a-vba-project/312a40f4-76a0-4b15-93f8-ea241f25ef61?page=1
