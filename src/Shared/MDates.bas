Attribute VB_Name = "MDates"
Option Explicit

' Copyright (c) 2017 Jason Peter Brown <jason@bitspaces.com>
'
' MIT License
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Public Function dateGmtToCookieDate(ByVal p_GmtDate As Date) As String
   dateGmtToCookieDate = "Wdy, DD Mon YYYY HH:MM:SS GMT"
   
   Mid$(dateGmtToCookieDate, 1, 3) = CookieWeekday(Weekday(p_GmtDate, vbSunday))
   Mid$(dateGmtToCookieDate, 6, 2) = Format$(Day(p_GmtDate), "00")
   Mid$(dateGmtToCookieDate, 9, 3) = CookieMonth(Month(p_GmtDate))
   Mid$(dateGmtToCookieDate, 13, 4) = Format$(Year(p_GmtDate), "0000")
   Mid$(dateGmtToCookieDate, 18, 2) = Format$(Hour(p_GmtDate), "00")
   Mid$(dateGmtToCookieDate, 21, 2) = Format$(Minute(p_GmtDate), "00")
   Mid$(dateGmtToCookieDate, 24, 2) = Format$(Second(p_GmtDate), "00")
End Function

Public Function dateVbLocalDateTimeToIso8601Utc(ByVal p_LocalDate As Date) As String
   Dim lt_LocalSystemTime As SYSTEMTIME
   Dim lt_UtcSystemTime As SYSTEMTIME
   
   With lt_LocalSystemTime
      .wYear = Year(p_LocalDate)
      .wMonth = Month(p_LocalDate)
      .wDay = Day(p_LocalDate)
      .wHour = Hour(p_LocalDate)
      .wMinute = Minute(p_LocalDate)
      .wSecond = Second(p_LocalDate)
   End With
   
   apiTzSpecificLocalTimeToSystemTime 0, VarPtr(lt_LocalSystemTime), VarPtr(lt_UtcSystemTime)
   
   dateVbLocalDateTimeToIso8601Utc = "0000-00-00T00:00:00Z"
   Mid$(dateVbLocalDateTimeToIso8601Utc, 1, 4) = Format$(lt_UtcSystemTime.wYear, "0000")
   Mid$(dateVbLocalDateTimeToIso8601Utc, 6, 2) = Format$(lt_UtcSystemTime.wMonth, "00")
   Mid$(dateVbLocalDateTimeToIso8601Utc, 9, 2) = Format$(lt_UtcSystemTime.wDay, "00")
   Mid$(dateVbLocalDateTimeToIso8601Utc, 12, 2) = Format$(lt_UtcSystemTime.wHour, "00")
   Mid$(dateVbLocalDateTimeToIso8601Utc, 15, 2) = Format$(lt_UtcSystemTime.wMinute, "00")
   Mid$(dateVbLocalDateTimeToIso8601Utc, 18, 2) = Format$(lt_UtcSystemTime.wSecond, "00")
End Function

Private Function CookieMonth(ByVal p_OneBasedMonthIndex As Long) As String
   ' Return English abbreviation for month
   Select Case p_OneBasedMonthIndex
   Case 1
      CookieMonth = "Jan"
   Case 2
      CookieMonth = "Feb"
   Case 3
      CookieMonth = "Mar"
   Case 4
      CookieMonth = "Apr"
   Case 5
      CookieMonth = "May"
   Case 6
      CookieMonth = "Jun"
   Case 7
      CookieMonth = "Jul"
   Case 8
      CookieMonth = "Aug"
   Case 9
      CookieMonth = "Sep"
   Case 10
      CookieMonth = "Oct"
   Case 11
      CookieMonth = "Nov"
   Case 12
      CookieMonth = "Dec"
   Case Else
      Err.Raise 5, , "Invalid one-based month index: " & p_OneBasedMonthIndex
   End Select
End Function

Private Function CookieWeekday(ByVal p_Weekday As VbDayOfWeek) As String
   ' Return English abbreviation for day of week
   Select Case p_Weekday
   Case vbSunday
      CookieWeekday = "Sun"
   Case vbMonday
      CookieWeekday = "Mon"
   Case vbTuesday
      CookieWeekday = "Tue"
   Case vbWednesday
      CookieWeekday = "Wed"
   Case vbThursday
      CookieWeekday = "Thu"
   Case vbFriday
      CookieWeekday = "Fri"
   Case vbSaturday
      CookieWeekday = "Sat"
   Case Else
      Err.Raise 5, , "Invalid day index: " & p_Weekday
   End Select
End Function

