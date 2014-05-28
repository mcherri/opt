' Copyright 2014 Moustapha Cherri

' This file is part of OPT (Outlook Prayer Times).

' OPT is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as
' published by the Free Software Foundation, either version 3 of
' the License, or (at your option) any later version.

' OPT is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.

' You should have received a copy of the GNU Lesser General Public License
' along with Foobar.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit
' Edit these to reflect your country
Public Const TimeZone As Double = 3#
Public Const DayLightSaving As Double = 0
Public Const Latitude As Double = 30.0566
Public Const Longitude As Double = 31.2262

' Don't edit these
Public Const PI As Double = 3.14159265358979

Public PMap(0 To 3) As Integer

#If VBA7 And Win64 Then
    Public TimerID As LongLong 'Need a timer ID to eventually turn off the timer. If the timer ID <> 0 then the timer is running
#Else
    Public TimerID As Long 'Need a timer ID to eventually turn off the timer. If the timer ID <> 0 then the timer is running
#End If

#If VBA7 And Win64 Then
    ' 64-bit Office
    Public Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal HWnd As LongLong, ByVal nIDEvent As LongLong, _
        ByVal uElapse As LongLong, _
        ByVal lpTimerFunc As LongLong) As LongLong
    Public Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal HWnd As LongLong, _
        ByVal nIDEvent As LongLong) As LongLong
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongLong
#Else '32-bit Office
    Public Declare Function SetTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" ( _
        ByVal HWnd As Long, _
        ByVal nIDEvent As Long) As Long
    Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If

#If VBA7 And Win64 Then ' 64-bit Office
Sub APITimerProc(ByVal HWnd As LongLong, ByVal uMsg As LongLong, _
        ByVal nIDEvent As LongLong, ByVal dwTimer As LongLong)
    ' Callback from Windows when SetTimer timer "pops"
    'KillTimer 0&, nIDEvent
    Call TriggerTimer
End Sub
#Else ' 32-bit Office
Sub APITimerProc(ByVal HWnd As Long, ByVal uMsg As Long, _
        ByVal nIDEvent As Long, ByVal dwTimer As Long)
    ' Callback from Windows when SetTimer timer "pops"
    'KillTimer 0&, nIDEvent
    Call TriggerTimer
End Sub
#End If

Public Sub ActivateTimer(ByVal nMinutes As Long)
    nMinutes = nMinutes * 1000 * 60 'SetTimer accepts milliseconds- convert to minutes.
    If TimerID <> 0 Then Call DeactivateTimer 'make sure there isn't already a timer.
    TimerID = SetTimer(0, 0, nMinutes, AddressOf APITimerProc)
    If TimerID = 0 Then
        MsgBox "Timer failure."
    End If
End Sub

Public Sub DeactivateTimer()
    #If VBA7 And Win64 Then
        Dim lSuccess As LongLong
    #Else
        Dim lSuccess As Long
    #End If
    
    lSuccess = KillTimer(0, TimerID)
 
End Sub

Public Sub TriggerTimer()
    Call DoIt
End Sub

Private Sub Init_PMap()
    PMap(0) = 2
    PMap(1) = 3
    PMap(2) = 5
    PMap(3) = 6
End Sub

Public Sub DoIt()
    Call RemoveAll
    Call Fill
End Sub

Public Sub Fill()
    Dim P As New PrayerTime
    P.CalcMethod = P.Egypt

    Call Init_PMap
    
    Dim N As Date, D As Date, D2 As Date
    N = Date
    
    Dim I As Integer
    For I = 0 To 27
        D = DateAdd("d", I, N)
        Dim T() As Double
        T = P.PrayerTimes(D)
        Dim J As Integer
        For J = 0 To 3
            Dim Time As Double, H As Integer, M As Integer
            Time = T(PMap(J))
            H = Int(Time)
            M = Int((Time - H) * 60 + 0.5)
            D2 = DateAdd("h", H, D)
            D2 = DateAdd("n", M, D2)
            Call Reserve(D2, GetPrayerName(J))
        Next
    Next
    

End Sub

Function GetPrayerName(I As Integer)
    Dim P As New PrayerTime
    P.CalcMethod = P.Egypt

    Call Init_PMap
    Dim Names() As String
    Names = P.TimeNames

    GetPrayerName = Names(PMap(I)) & " Prayer"
End Function

Private Sub Reserve(D As Date, Subject As String)
    Dim Item As Object
  
    Set Item = Application.CreateItem(olAppointmentItem)
 
    Item.Subject = Subject
    Item.Start = D '#6/3/2013 1:30:00 PM#
    Item.Duration = 15
    Item.Categories = "Green Category"
    Item.Save

End Sub

Public Sub RemoveAll()
    Dim oCalendar As Outlook.Folder
    Dim oItems As Outlook.Items
    Dim oResItems As Outlook.Items
    Dim oAppt As Outlook.AppointmentItem
    Dim strRestriction As String, strOr As String
    
    Set oCalendar = Application.Session.GetDefaultFolder(olFolderCalendar)
    Set oItems = oCalendar.Items
    
    strRestriction = ""
    strOr = ""
    
    Dim I As Integer
    For I = LBound(PMap) To UBound(PMap)
        strRestriction = strRestriction & strOr & "[Subject] = '" & GetPrayerName(I) & "'"
        strOr = " OR "
    Next
    
    strRestriction = "(" & strRestriction & ") AND [Categories] = 'Green Category'"
    
    Set oResItems = oItems.Restrict(strRestriction)
    oResItems.Sort "[Start]"
    
    For I = oResItems.Count To 1 Step -1 'Iterates from the end backwards
        Set oAppt = oResItems.Item(I)
        oAppt.Delete
    Next
End Sub

Public Sub Remove(Item As Object)
    
    Dim P As New PrayerTime
    P.CalcMethod = P.Egypt

    Call Init_PMap
    Dim Names() As String
    Names = P.TimeNames
    
    Dim I As Integer
    For I = 0 To 3
        If Names(PMap(I)) & " Prayer" = Item.Subject And Item.Categories = "Green Category" Then
            Item.Delete
            Exit For
        End If
    Next
End Sub

