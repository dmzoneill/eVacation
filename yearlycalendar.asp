<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!-- #include virtual="/eVacation/common/objects/calendar.asp" -->
<!--#include virtual="/eVacation/common/calendarfunctions.asp" -->

<%
	
	strCurrentPageName = "Employee Leave Calendar"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView

	If not objCurrentUser.IsEELeaveTracked then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If

	mWriteHMTLTop strCurrentPageName
	
	
		Dim myRows
		Dim myPrintRows
		Dim myTotalRows
		Dim myColumns
		Dim myMonth
		Dim myYear
		Dim myMonthsToDisplay  
        Dim myMonthsCheck
		Dim myMonthsDisplayed
		Dim MyCalendar

        On Error Resume Next
		
		'***CHECKS IF FORM HAS BEEN SUBMITTED WITH DATES AND MONTHS NEEDED TO PRINT CALENDAR
		if Request.Form("startMonth") = "" or Request.Form("yearOfCal") = "" or Request.Form("monthsToDisplay") = "" then
			'**IF NOT SET THEN CHOOSE TO DISPLAY FULL YEAR
			myMonthsDisplayed = 1
			
			myMonth = 1
			myYear = Year(Date())
			myMonthsToDisplay = 12
		else
			myMonthsDisplayed = 1
		    if IsNumeric(Request.Form("startMonth")) and IsNumberic(Request.Form("yearOfCal")) and IsNumeric(Request.Form("monthsToDisplay")) then
                myMonth = CInt(Request.Form("startMonth"))
			    myYear = CInt(Request.Form("yearOfCal"))
		        myMonthsToDisplay = CInt(Request.Form("monthsToDisplay"))
            else
                response.redirect CONST_APPLICATION_PATH & "/caldisplay.asp"
            end if
		end if
		
		'**CALCULATING THE TOTAL NUMBER OF ROWS REQUIRED FOR THE MONTHS TO BE DISPLAYED
		myMonthsCheck = myMonthsToDisplay
		myTotalRows = 0
		
		do while myMonthsCheck > 0
		
			myTotalRows = myTotalRows + 1
			myMonthsCheck = myMonthsCheck - 6
			
		loop
			
			
		'***DECIDES ON THE TOTAL NUMBER OF ROWS TO DISPLAY, DEPENDING ON THE AMOUNT OF MONTHS COULD BE LARGER
		for myRows = 1 to myTotalRows
		%>
		<table cellspacing="0" cellpadding="0"">
		<%
		for myPrintRows = 1 to 3
		%>
			<tr>
		<%
			'**SETS HOW MANY MONTHS ARE DISPLAYED IN EACH ROW
			for myColumns = 1 to 2
			%>
				<td valign="top">
				<%
					'**CHECKS IF ALL THE MONTHS THAT NEED TO BE DISPLAYED HAVE ALREADY BEEN DISPLAYED
					if myMonthsDisplayed > myMonthsToDisplay then
									
					else
						'**SETS THE NEW CALENDAR OBJECT TO THE VARIABLE
						Set MyCalendar = New Calendar
						
						'**SETS THE MONTH AND YEAR TO BE DISPLAYED IN THE CALENDAR
						MyCalendar.setDate(myMonth & "/7/" & myYear)
						
						'**INCREMENETS THE MYMONTH VARIABLE BY ONE FOR THE NEXT MONTH TO BE DISPLAYED
						myMonth = myMonth + 1
						
						'**CHECKS THAT THE MONTHS HASN'T RAN OVER THE YEAR END, IF IT HAS INCREMENTS YEAR
						if myMonth > 12 then
							myMonth = 1
							myYear = myYear + 1
						end if
						
						'SETS DIFFERENT VARIABLES OF THE CALENDAR OBJECT
						MyCalendar.Width = 340
						MyCalendar.SetDrawStyle true
						
						'**ADDS ALL THE LEAVES TO THE HOLIDAY OBJECT
						mAddCalendarHolidays objEEtoView
						
						'**REMOVES THE DATE SELECT FEATURE ON THE DISPLAY OF THE CALENDARS
						MyCalendar.ShowDateSelect = false
						
						'**DISPLAYS THE CALENDAR AND THEN INCREMEMTS THE NUMBER OF MONTHS ALREADY DISPLAYED BY 1
						MyCalendar.Draw()
						myMonthsDisplayed = myMonthsDisplayed + 1
					end if
				%>
				</td>
			<%
			next
			%>
			</tr>
			<%
			next
			%>
			</table>
			<% if myPrintRows = 4 then 
			%>
	
	<br style="page-break-after: always;">
	<%
	myPrintRows = myPrintRows + 1
	else
	end if
		next
		%>
		<table class="noDisplay">
	<tr>
		<td colspan="6" align="center">
			<input type="button" value="Print" onclick="window.print()">
		</td>
	</tr>
	</table>

<%
	
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
