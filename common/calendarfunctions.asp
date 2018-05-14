<%
    '***** FUNCTIONS *****
    'mAddCalendarHolidays	ByRef objEEtoView
    
    
    function mAddCalendarHolidays(ByRef objEEtoView)
    'create a new connection because need to modify cursor location(error otherwise)
			
			Dim rstLeaveHols
			Dim cmGetEmployeeHolidayData
			Dim m_cnDB
			
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
			cmGetEmployeeHolidayData.Parameters.Append cmGetEmployeeHolidayData.CreateParameter("@vWWID", adChar, adParamInput, 8, objEEtoView.WWID)
			Set rstLeaveHols = cmGetEmployeeHolidayData.Execute
			
			dim myArrColors(20)
			dim myCount
			
			myCount = 0
		
		        myArrColors(1) = "0000FF"           'Change the color to blue for outpu to calendar (Chris)
				myArrColors(2) = "0000FF"
				myArrColors(3) = "0000FF"
				myArrColors(4) = "0000FF"
				myArrColors(5) = "0000FF"
				myArrColors(6) = "0000FF"
				myArrColors(7) = "0000FF"
				myArrColors(8) = "0000FF"
				myArrColors(9) = "0000FF"
				myArrColors(10) = "0000FF"
	        	
	
				'myArrColors(1) = "990000"
				'myArrColors(2) = "000099"
				'myArrColors(3) = "009900"
				'myArrColors(4) = "990033"
				'myArrColors(5) = "FF0099"
				'myArrColors(6) = "FF9933"
				'myArrColors(7) = "3090F0"
				'myArrColors(8) = "90f030"
				'myArrColors(9) = "CC66FF"
				'myArrColors(10) = "FF9933"
			
			'***RUN THROUGH THE FULL LIST OF LEAVE RECORDS
			do while not rstLeaveHols.eof
				dim counter
				dim i
				dim myNewColor
				dim firstInitial
				dim secondInitial

				
				
				myNewColor = trim(Mid(rstLeaveHols.fields.item(4).value,1,6))
				
				
				'***RUN THROUGH DATES OF LEAVES
				for i = rstLeaveHols.fields.item(2).value to rstLeaveHols.fields.item(3).value
					
					
					dim dayOfHol
					dim monOfHol
					dim yearOfHol
					dim mySize
					
					dim j
					
					dayOfHol = Day(i)
					monOfHol = Month(i)
					yearOfHol = Year(i)
					mySize = "300"
					
					
					'***CHECK THAT THE DATE IS FOR THE DAY/MONTH/YEAR BEING DISPLAYED AND THEN ADD
					if Year(MyCalendar.GetDate()) = yearOfHol AND Month(MyCalendar.GetDate()) = monOfHol AND not rstLeaveHols.fields.item(5).value = "" AND not  WeekDay(i)=1 AND not WeekDay(i)=7 then
							if not rstLeaveHols.fields.item(6).value = "" then
								
							else
								MyCalendar.Days(dayOfHol).AddActivity1 (Mid(rstLeaveHols.fields.item(0).value,1,1)) & "." & (Mid(rstLeaveHols.fields.item(1).value,1,1)), myArrColors(myCount), rstLeaveHols.fields.item(4).value, mySize, trim(rstLeaveHols.fields.item(0).value) & " " & trim(rstLeaveHols.fields.item(1).value)
							end if
					end if
				next
				
				if myCount <> 11 then
					myCount = myCount + 1
				else
					myCount = 0
				end if
				
				
				'***MOVE TO NEXT RECORD
				rstLeaveHols.movenext
			loop
			
	end function
	

%>
