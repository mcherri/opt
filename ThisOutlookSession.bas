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

Dim ItemEvnts As ItemEvents

Private Sub Application_Startup()
    Set ItemEvnts = New ItemEvents
    Call FillCalendar.ActivateTimer(7 * 24 * 60)
    Call FillCalendar.DoIt
End Sub
