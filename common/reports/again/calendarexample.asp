<% Option Explicit %>
<!--#include virtual="/eVacation_DEV/common/appglobal.asp" -->
<!-- #include file="calendar.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY LINK="blue" ALINK="blue" VLINK="blue">
<%
	Private rstPubHols
	
	Dim MyCalendar
	Dim daysInMonth1
	Dim dayOfWeek
	Dim ThisYear
	Dim cmGetEmployeeHolidayData
	Dim rstLeavePeriodsInYear
	Dim m_cnDB
	Dim manID
	
	manID = 10705764
	ThisYear = 2007
	dayOfWeek = WeekDay(mFirstDayOfYear(ThisYear))
	daysInMonth1 = mGetDaysInMonthForDate(mFirstDayOfYear(ThisYear))
	
	'create a new connection because need to modify cursor location(error otherwise)
	set m_cnDB = Server.CreateObject("ADODB.Connection")
	m_cnDB.ConnectionString = CONST_ADO_EVACATION_CONNECTION_STRING	
	m_cnDB.CursorLocation = adUseClient
	m_cnDB.Open
	
	'*********************************************
	'**				
	'** retrieve the employees leave for the year
	'**
	'*********************************************
	Set cmGetEmployeeHolidayData = Server.CreateObject("ADODB.Command")
	Set cmGetEmployeeHolidayData.ActiveConnection =  m_cnDB
	cmGetEmployeeHolidayData.CommandType = 4
	cmGetEmployeeHolidayData.CommandText = "dbo.hol_cal_display"
	Set rstPubHols = cmGetEmployeeHolidayData.Execute

	' Create the calendar
	Set MyCalendar = New Calendar
	
	' Set the visual properties
	MyCalendar.Top = 0 'Sets the top position
	MyCalendar.Left = 0 'Sets the left position
	MyCalendar.Position = "absolute" 'Relative or Absolute positioning
	MyCalendar.Height = "600" 'Sets the height
	MyCalendar.Width = "600" 'Sets the width
	MyCalendar.TitlebarColor = "darkblue" 'Sets the color of the titlebar
	MyCalendar.TitlebarFont = "arial" 'Sets the font face of the titlebar
	MyCalendar.TitlebarFontColor = "white" 'Sets the font color of the titlebar
	MyCalendar.TodayBGColor = "skyblue" 'Sets the highlight color of the current day
	MyCalendar.ShowDateSelect = True 'Toggles the Date Selection form.
	
	' Add event code for when a day is clicked on. Notice
	' that when run inside your browser, "$date" is replaced
	' by the date you click on. 
	MyCalendar.OnDayClick = "javascript:alert('You clicked on this date: $date')"
	
	dim myColors
	dim myColorsNum
	
	myColors = Array("yellow","green","pink","blue","orange","gold")
	myColorsNum = 0
	
	do while not rstPubHols.eof
		dim counter
		'for counter=0 to rstPubHols.fields.count-1
			dim i
			'i = rstPubHols.fields.item(3).value
			for i = rstPubHols.fields.item(2).value to rstPubHols.fields.item(3).value
			dim dayOfHol
			dim monOfHol
			dim yearOfHol
			dim myNewColor
			
			
			dayOfHol = Day(i)
			monOfHol = Month(i)
			yearOfHol = Year(i)
			myNewColor = "#" & Mid(rstPubHols.fields.item(4).value,3,7) & "F"
			
			if Year(MyCalendar.GetDate()) = yearOfHol then
				if Month(MyCalendar.GetDate()) = monOfHol then
					if not rstPubHols.fields.item(5).value = "" then
                        if not rstPubHols.fields.item(7).value = 0 then
					        MyCalendar.Days(dayOfHol).AddActivity rstPubHols.fields.item(0).value & rstPubHols.fields.item(1).value, myNewColor
                        end if
					end if
				end if
			end if
			next
			if myColorsNum = 5 then
							myColorsNum = 0
						else
							myColorsNum = myColorsNum + 1
						end if
		'next
		rstPubHols.movenext
	loop
	
	' Add holidays to the calendar
	Select Case Month(MyCalendar.GetDate())
		' January
		Case 1
			' Add New Years Day
			MyCalendar.Days(1).AddActivity "<small><b>New Years</b></small>", "pink"
		
		' December
		Case 12
			' Add Christmas Day
			MyCalendar.Days(25).AddActivity "<small><b>Christmas</b></small>", "pink"
	End Select
	
	' Draw the calendar to the browser
	MyCalendar.Draw()
%>
</BODY>
</HTML>
