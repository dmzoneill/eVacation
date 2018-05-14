<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
Function LeaveDaysInPeriod (ByVal locdatPeriodStartDate, ByVal locstrPeriodStartTime,Byval locdatPeriodEndDate, byval locstrPeriodEndTime, ByVal locdatReportStartDate, ByVal locdatReportEndDate)

			Dim locdblDays
			Dim blnStartDateChange
			Dim blnEndDateChange
			
			blnStartDateChange = False
			blnEndDateChange = False
			
			If locdatPeriodStartDate > locdatPeriodEndDate or locdatPeriodEndDate < locdatPeriodStartDate then
				LeaveDaysInPeriod = 0
				Exit Function
			End If
			
			If locdatPeriodEndDate > locdatReportEndDate then
				 locdatPeriodEndDate = locdatReportEndDate
				 blnEndDateChange = true
			End If
			'CA added parameter locdatRptStartDate
			If (locdatPeriodStartDate < locdatReportStartDate) and (locdatPeriodEndDate >= locdatReportStartDate) then
				 locdatPeriodStartDate = locdatReportStartDate
				 blnStartDateChange = true
			End If
			
												
			locdblDays = DateDiff("y",locdatPeriodStartDate,locdatPeriodEndDate) + 1

			'*** If the start date is in the requested period, and is PM and not a weekend day 
			'or public holiday, take off half a day.
			If locstrPeriodStartTime = "PM" then
				if (not glbPublicHolidays.Contains(locdatPeriodStartDate, False)) and (not mIsWeekendDay(locdatPeriodStartDate)) then
					if not blnStartDateChange then
						locdblDays = locdblDays - 0.5
					end if
				end if
			end if
			
			If locstrPeriodEndTime = "AM" then				
				if (not glbPublicHolidays.Contains(locdatPeriodEndDate, False)) and (not mIsWeekendDay(locdatPeriodEndDate)) then
					if not blnEndDateChange then
						locdblDays = locdblDays - 0.5
					end if
				end if
			end if

					
			LeaveDaysInPeriod = locdblDays - glbPublicHolidays.CountInPeriod(locdatPeriodStartDate,locdatPeriodEndDate,False) - mCountWeekendDays(locdatPeriodStartDate, locdatPeriodEndDate)
			
end function
%>




<%
	Server.ScriptTimeOut = 6000	'10 Minutes (6000 seconds)
	
	Dim locCmd
	Dim locRS
				
	Dim fldstrWWID
	Dim fldstrFullName
	dim fldstrNextLevelNm' added MFILLAST
	dim fldstrOrgUnit' added MFILLAST
	Dim fldstrLeaveType	
	Dim flddatStartDate
	dim fldstrStartTime
	Dim flddatEndDate
	dim fldstrEndTime
	Dim flddatLeaveRequestRaised
	Dim flddatApproved
	Dim flddatRejected 
	Dim flddatCancelRequested
	Dim flddatCancelApproved
	Dim flddatCancelRejected
	Dim locArrColumnWidths(6)
	Dim locArrColumnHdrBGColours(6)
	Dim locArrColumnNames(6)
	Dim locdatRptStartDate  'Report Input Start Date
	Dim locdatRptEndDate    'Report Input End Date
	Dim locstrErrorMessage
	dim locdatRptStartDateForStoredProcedure  'Report Input Start Date For Stored Procedure
	dim loclngRptStartYear 'the year of Report Input Start Date
	
	dim strEmployeeStatus
	dim strSite
	
	dim blnShow
	
	strCurrentPageName = "Payroll Report Lite"

	locdatRptStartDate = Request.querystring("datStartDate") 'CA this is the real date that the user enters
			
	locdatRptEndDate = Request.querystring("datEndDate")
	
	strEmployeeStatus = Request.QueryString("status")
	strSite = Request.QueryString("site")
	
    mDebugPrint locdatRptStartDate
    mDebugPrint locdatRptEndDate
    mDebugPrint strEmployeeStatus
    mDebugPrint strSite

	locstrErrorMessage = ""
	
	if not isdate(locdatRptStartDate) then
		locstrErrorMessage = locstrErrorMessage & "<br> - The start date entered is invalid."
	end if
	
	if not isdate(locdatRptEndDate) then
		locstrErrorMessage = locstrErrorMessage & "<br> - The end date entered is invalid."
	end if
	
	if isdate(locdatRptStartDate) and isdate(locdatRptEndDate) then
		locdatRptStartDate = cDate(locdatRptStartDate)
		locdatRptEndDate = cDate(locdatRptEndDate)
		if datepart("yyyy",locdatRptStartDate) <> datepart("yyyy",locdatRptEndDate) then
			locstrErrorMessage = locstrErrorMessage & "<br> - The start and end dates must be in the same year."
		end if
		
		if locdatRptStartDate > locdatRptEndDate then
			locstrErrorMessage = locstrErrorMessage & "<br> - The start date must be on or before the end date."
		end if
	end if
	
	if locstrErrorMessage <> "" then
		mWriteHMTLTop strCurrentPageName
		mWriteGeneralError "Sorry - the form was not correctly filled in...<br>" & locstrErrorMessage & "<br><br>Please close this window and try again.<br>", False
		mWritePageFooter
	else

		loclngRptStartYear = DatePart("yyyy", locdatRptStartDate) - 1
		'the system starts on 2001
		'locdatRptStartDateForStoredProcedure =  mFirstDayOfYear(loclngRptStartYear)
		     'CA this locdatRptStartDateForStoredProcedure is a fix to get data from Database 
		     'for 2 years (1year from 1 Jan) in advance to catch Leave Period which which 
		     'overlaps the Report Requested Dates).  Leave Period which span over a large period
		     'is ELP (average 2 months).


	    '***** Set content-type to Excel (for display in browser as Excel workbook). *******
	'comment this line to display errors
	    Response.ContentType = "application/vnd.ms-excel" 
	
		'***** SET THIS HTML DOCUMENT TO BE RECOGNISED BY THE BROWSER AS AN EXCEL SPREADSHEET ******
		response.write "<HTML xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
		
			'***** SET UP PAGE ****
			response.write "<HEAD>"
				response.write "<style>"
				  	response.write "<!--table"
					  	response.write "@page"
				     	response.write "{mso-header-data:""&Ce-Vacation - Monthly Payroll Absences Report\000A"
				     	Response.Write "Employee with Leave Period/s within\000A"'CA
				     	response.write "Report Start Date: " & mFormatDate(locdatRptStartDate,"medium") & " - Report End Date: " & mFormatDate(locdatRptEndDate,"medium") & "\000A"
				     	response.write "Print Date\: &D" & "    Page &P"";"
						response.write "mso-page-orientation:landscape;}"
				     	response.write "br"
				     	response.write "{mso-data-placement:same-cell;}"
				  	response.write "-->"
				response.write "</style>"
				
				'*** WRITE OUT THE EXCEL PRINTER OPTIONS ***
			  	response.write "<!--[if gte mso 9]>"
			  		response.write "<xml>"
					   	response.write "<x:ExcelWorkbook>"
					    	response.write "<x:ExcelWorksheets>"
					     	response.write "<x:ExcelWorksheet>"
					      	response.write "<x:Name>Monthly Payroll Absences Report</x:Name>"
					      	response.write "<x:WorksheetOptions>"
					       	response.write "<x:FitToPage/>"
					       	response.write "<x:Print>"
					        	response.write "<x:RowColHeadings/>"
								response.write "<x:FitWidth>1</x:FitWidth>"
								response.write "<x:FitHeight>80</x:FitHeight>"
					        	response.write "<x:ValidPrinterInfo/>"
					        	response.write "<x:Gridlines/>"
					       	response.write "</x:Print>"
					      	response.write "</x:WorksheetOptions>"
					     	response.write "</x:ExcelWorksheet>"
					    	response.write "</x:ExcelWorksheets>"
					   	response.write "</x:ExcelWorkbook>"
				  	response.write "</xml>"
				response.write "<![endif]-->"
			response.write "</HEAD>"
			response.write "<BODY>"

			'** Get the Data from Stored Procedure usp_ee_payrollreportlite
			Set locCmd = Server.CreateObject("ADODB.Command")
	
			Set locCmd.ActiveConnection = glbConnection
			locCmd.CommandType = adCmdStoredProc

			blnShow = false
			
			if trim(LCase(strEmployeeStatus)) <> "all" then
				locCmd.CommandText = "pr_evc_payroll_report_filtered_without_site"
				locCmd.Parameters.Append locCmd.CreateParameter("@vStartDate", adDBTimeStamp, adParamInput, , locdatRptStartDate) 'CA
				locCmd.Parameters.Append locCmd.CreateParameter("@vEndDate", adDBTimeStamp, adParamInput, , locdatRptEndDate)						
				locCmd.Parameters.Append locCmd.CreateParameter("@vStatus", adVarChar, adParamInput, 20, trim(strEmployeeStatus)) 'CA
				if trim(LCase(strSite)) <> "all" then
					locCmd.CommandText = "pr_evc_payroll_report_filtered_with_site"
					locCmd.Parameters.Append locCmd.CreateParameter("@vSite", adVarChar, adParamInput, 3, trim(strSite))
				end if
				blnShow = True
			elseif trim(LCase(strSite)) <> "all" then
				locCmd.CommandText = "dbo.pr_evc_payroll_report_filtered_site_only"
				locCmd.Parameters.Append locCmd.CreateParameter("@vStartDate", adDBTimeStamp, adParamInput, , locdatRptStartDate) 'CA
				locCmd.Parameters.Append locCmd.CreateParameter("@vEndDate", adDBTimeStamp, adParamInput, , locdatRptEndDate)
				locCmd.Parameters.Append locCmd.CreateParameter("@vSite", adVarChar, adParamInput, 3, trim(strSite))						
				blnShow = True
			end if
			
			'if not blnShow then		'[MFILLAST 08-2006] comment : seems to be not working
			'	locCmd.CommandText = "usp_ee_payrollreportlite2"			
			'	locCmd.Parameters.Append locCmd.CreateParameter("datRptStartDate", adDBTimeStamp, adParamInput, , locdatRptStartDate) 'CA
			'	locCmd.Parameters.Append locCmd.CreateParameter("datRptEndDate", adDBTimeStamp, adParamInput, , locdatRptEndDate)
			'end if
			mDebugPrint locCmd.CommandText	
			Set locRS = locCmd.Execute

			'WWID, FullName, LeavePeriodId, LeaveType, StartDate, StartTime, EndDate ,EndTime

				Set fldstrWWID = locRS("WWID")
				Set fldstrFullName = locRS("FullName")
				Set fldstrNextLevelNm = locRS("NextLevelNm")' added MFILLAST
				Set fldstrOrgUnit = locRS("DepartmentNm")' added MFILLAST
				Set fldstrLeaveType = locRS("LeaveType")
				Set flddatStartDate = locRS("StartDate")
				Set fldstrStartTime = locRS("StartTime")
				Set flddatEndDate = locRS("EndDate")
				Set fldstrEndTime = locRS("EndTime")
				set flddatLeaveRequestRaised = locRS("LeaveRequestRaised")
				set flddatApproved = locRS("Approved")
				set flddatRejected = locRS("Rejected")
				set flddatCancelRequested = locRS("CancelRequested")
				set flddatCancelApproved = locRS("CancelApproved")
				set flddatCancelRejected = locRS("CancelRejected")
				
				'** Got the Data from Stored Procedure usp_ee_payrollreportlite
				
				'** Display Header **
				response.write "<TABLE>"
				response.write "<TR>"
					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "WWID" 'CA
					response.write "</TD>"

					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "Last Name, First Name" 'CA
					response.write "</TD>"

					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "Manager Name" 'CA
					response.write "</TD>"
					
					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "Org Unit" 'CA
					response.write "</TD>"

					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "Type of absence"
					response.write "</TD>"

					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "No. of days of absence"
					response.write "</TD>"

					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "1st day of absence"
					response.write "</TD>"

					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "Last day of absence"
					response.write "</TD>"	
					
					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "Date Raised"
					response.write "</TD>"

					response.write "<TD BGColor=""#c0c0c0"">"
						Response.Write "Date Approved"
					response.write "</TD>"
									
				response.write "</TR>"
				'** Finish Display Header **
				
				'** Display Records **
				Do Until locRS.EOF
			
					'This record set returns EE with no leave period
					'and ee with leave period with Leave Period Start Date >= Report Start Date.
					'Now check if this leave period start date is in the same year as the year for
					'report start date.  We are only interested in the Leave Period start date =< Report End date
					'OR EE with no leave period (the GUI will not let a user enter different year for 
					'RptStartDate and RptendDate.)
					'If	IsNull(flddatStartDate) OR (flddatStartDate <= locdatRptEndDate)Then
						'CA - it is important the SP is order by WWID, then Start Date
						'for this record if the Record's End date ends after the User input 
						'Report Start Date then show it.
					
					IF (IsNull(flddatCancelApproved) or not(flddatCancelApproved <> "")) and (IsNull(flddatRejected) OR not (flddatRejected <> "" )) then  'added by [MFILLAST 08-2006] to not display cancelled or rejected requests
								
							response.write "<TR>"
								response.write "<TD>"
								response.write fldstrWWID
								response.write "</TD>"
								
								response.write "<TD>"
								response.write fldstrFullName
								response.write "</TD>"
								
								response.write "<TD>"
								response.write fldstrNextLevelNm
								response.write "</TD>"								
								
								response.write "<TD>"
								response.write fldstrOrgUnit
								response.write "</TD>"
								
								response.write "<TD>"
								response.write fldstrLeaveType	
								response.write "</TD>"
								
								response.write "<TD>"
								response.write LeaveDaysInPeriod(flddatStartDate, fldstrStartTime, flddatEndDate, fldstrEndTime, locdatRptStartDate, locdatRptEndDate)
								response.write "</TD>"			
								
								response.write "<TD>"
								response.write mFormatDate(flddatStartDate,"medium with day") & " " & fldstrStartTime
								response.write "</TD>"
								
								response.write "<TD>"
								response.write mFormatDate(flddatEndDate,"medium with day") & " " & fldstrEndTime
								response.write "</TD>"
								
								response.write "<TD>"
								response.write mFormatDate(flddatLeaveRequestRaised,"medium with time")
								response.write "</TD>"
								
								response.write "<TD>"
								response.write mFormatDate(flddatApproved,"medium with time") 
								response.write "</TD>"
								
							response.write "</TR>"
						end if
						
					
					locRS.MoveNext
				Loop
				'** Finish Display Records **

				
				Set fldstrWWID = Nothing
				Set fldstrFullName = Nothing
				Set fldstrLeaveType = Nothing
				Set flddatStartDate = Nothing
				Set fldstrStartTime = Nothing
				Set flddatEndDate = Nothing
				Set fldstrEndTime = Nothing
				Set fldstrOrgUnit = Nothing

				locRS.Close
	
				Set locRS = nothing
				Set locCmd = nothing
				
				response.write "</TABLE>"
				
			response.write "</BODY>"
		response.write "</HTML>"	
	end if
%>
