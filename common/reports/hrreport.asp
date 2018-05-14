<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	Response.Expires = -1000
	Server.ScriptTimeout = 12000 
	'Response.Buffer = True
	
	'<!--#include virtual="/eVacation_DEV2/common/appglobal.asp" -->
	' *** Employee Object ***
	Class clsEmployeeHRData
		'=================================================================
        'Description:   Calculates and stores HR information of an Intel Employee.
		'Properties:                Type:   Perm:       Source:     Properties Required:
		'	WWID					str												
		'	EndDate					dte												
		'	Status					str												
		'	Exempt					str																						
		'	Name					str												
		'	Manager					str												
		'	Site					str												
		'	BalanceAtEndOfYear		int														
		'	NumberOfDaysWorked		int												
		'	EndOfYear				dte																						
		'	VacationTaken			int																						
		'	VacationApproved		int												
		'	OtherLeaveTaken			int														
		' 	CarryOver				int												
		'	PreviousCarryOver		int												
		'	Entitlement				int												
		' 	DepositELP				int												
		' 	ELPUsed					int												
		' 	OrgUnit					str												
		
		'** Private Properties
		Private strWWID
		Private strStatus
		Private strExempt
		Private strName
		Private strManagerName
		Private strSite
		Private dteOriginalStartDt
		Private dteLastDateOfHire
		Private intLeaveBalanceAtEndOfYear
		Private intNumberOfDaysWorked
		Private intEndOfYear
		Private intNumVacationApproved ' Added [MOF 11/25/08] 
		Private intNumVacationTaken
		Private intNumOtherLeaveTaken
		Private intCarryOverEOY
		
		Private dblBasicEntitlement
		Private dblCarryOver
		Private dblLegalAdjustemntAccrued
		
		Private dblTotalRestDays	
		Private dblTotalLeaveBooked
		
		Private dblELPBanked
		Private dblELPUsed 'added [MFILLAST 08-2006]
		
		Private intYear
		Private dteStartOfYear, dteEndOfYear
		Private dteSelectedStartDate, dteSelectedEndDate
		Private dteStartOfNextMonth					
					
		Private blnBlueBadge
		Private dteDOB
		Private dteEndDate 'added [MFILLAST 08-2006]
		Private strOrgUnit'added [MFILLAST 08-2006]
		Private rstLeavePeriodsInYear
		
		Private lngDaysInYear
		Private lngWeekendDaysInYear
		Private lngPublicHolidaysInYear
		Private lngWorkingDaysInYear
		
		
		'** Public Properties
		Public Property Get WWID
			WWID = strWWID
		End Property

		Public Property Get EndDate
			WWID = dteEndDate
		End Property

		Public Property Get Status
			Status = strStatus
		End Property
		
		Public Property Get Exempt
			Exempt = strExempt
		End Property
		
		Public Property Get Name
			Name = strName
		End Property
		
		Public Property Get Manager
			Manager = strManagerName 
		End Property
		
		Public Property Get Site
			Site = strSite
		End Property
		
		
		Public Property Get BalanceAtEndOfYear
			BalanceAtEndOfYear = mRoundToNextHalf(intLeaveBalanceAtEndOfYear)
		End Property
		
		Public Property Get NumberOfDaysWorked
			NumberOfDaysWorked = mRoundToNextHalf(intNumberOfDaysWorked)
		End Property
		
		
		Public Property Get EndOfYear
			EndOfYear = intEndOfYear
		End Property
		
		Public Property Get VacationTaken
			VacationTaken = intNumVacationTaken
		End Property
		
		Public Property Get VacationApproved 	' Added [MOF 11/25/08] 
			VacationApproved = intNumVacationApproved
		End Property
		
		Public Property Get OtherLeaveTaken
			OtherLeaveTaken = intNumOtherLeaveTaken
		End Property
		
		Public Property Get CarryOver
			CarryOver = mRoundToNextHalf(intCarryOverEOY)
		End Property
'[MFILLAST 08-2006]
		Public Property Get PreviousCarryOver
			PreviousCarryOver = mRoundToNextHalf(dblCarryOver)
		End Property
'[MFILLAST 08-2006]
		Public Property Get Entitlement
			Entitlement = mRoundToNextHalf(dblBasicEntitlement)
		End Property
		
'[MFILLAST 08-2006]
		Public Property Get DepositELP
			DepositELP = mRoundToNextHalf(dblELPBanked)
		End Property			
		
'[MFILLAST 08-2006]
		Public Property Get ELPUsed
			ELPUsed = mRoundToNextHalf(dblELPUsed)
		End Property	
		
'[MFILLAST 08-2006]
		Public Property Get OrgUnit
			OrgUnit = strOrgUnit
		End Property	
				
		'** Private Methods
		Private Function DaysConsecutive()
		
			Dim dblDays
			
			If (not isdate(rstLeavePeriodsInYear("datStartDate"))) or not isdate(rstLeavePeriodsInYear("datEndDate")) then
				DaysConsecutive = 0
				Exit Function
			End if
			
			dblDays = DateDiff("y",rstLeavePeriodsInYear("datStartDate"),rstLeavePeriodsInYear("datEndDate")) + 1

			If rstLeavePeriodsInYear("strStartTime") = "PM" then
				if (not glbPublicHolidays.Contains(rstLeavePeriodsInYear("datStartDate"), False)) and (not mIsWeekendDay(rstLeavePeriodsInYear("datStartDate"))) then
					dblDays = dblDays - 0.5
				end if
			end if
			
			if rstLeavePeriodsInYear("strEndTime") = "AM" then
				if (not glbPublicHolidays.Contains(rstLeavePeriodsInYear("datEndDate"), False)) and (not mIsWeekendDay(rstLeavePeriodsInYear("datEndDate"))) then
					dblDays = dblDays - 0.5
				end if
			end if

			DaysConsecutive = dblDays
			
		End Function
		
		Private Function Days()
			Dim dblDays
			
			If (not IsDate(rstLeavePeriodsInYear("datStartDate"))) or not isdate(rstLeavePeriodsInYear("datEndDate")) then
				Days = 0
				Exit Function
			End if

			dblDays = DaysConsecutive
			
			dblDays = dblDays - glbPublicHolidays.CountInPeriod(rstLeavePeriodsInYear("datStartDate"), rstLeavePeriodsInYear("datEndDate"),False)
			dblDays = dblDays - mCountWeekendDays(rstLeavePeriodsInYear("datStartDate"), rstLeavePeriodsInYear("datEndDate"))
			Days = dblDays
			
		End Function
		
		Private Function StopsLegalAdjAccrual()
			If (rstLeavePeriodsInYear("dblDaysBeforeStopsLegalAdjAccrual") >= 0) then
					If Days > rstLeavePeriodsInYear("dblDaysBeforeStopsLegalAdjAccrual") then
						StopsLegalAdjAccrual = True
					Else
						StopsLegalAdjAccrual = False
					End If
			Else
				StopsLegalAdjAccrual = False
			End If
		End Function
		
		Private Function LeaveDaysInPeriod(ByVal i_dteStart, ByVal i_dteEnd, ByVal blnAffectOnlyLegalAllowance, ByVal blnBalanceAffecting, ByVal blnELP)
		
			Dim dteStartDate
			Dim dteEndDate
			Dim dblDays
			
			Dim dblLeaveDaysInPeriod
			
			dblLeaveDaysInPeriod = 0
			
			if rstLeavePeriodsInYear.EOF then
				if not rstLeavePeriodsInYear.BOF then
					rstLeavePeriodsInYear.MoveFirst	
				end if
			else
				if not rstLeavePeriodsInYear.BOF then
					rstLeavePeriodsInYear.MoveFirst	
				end if
			end if
						
			while not rstLeavePeriodsInYear.EOF
							
				' Check if this leave is not confirmed
				If isdate(rstLeavePeriodsInYear("datConfirmed")) then
					intNumVacationTaken = intNumVacationTaken + 1
				End If
				
				If rstLeavePeriodsInYear("datStartDate") > i_dteEnd or rstLeavePeriodsInYear("datEndDate") < i_dteStart then
		
					dblDays = dblDays + 0
		
				else
				
					If rstLeavePeriodsInYear("datStartDate") < i_dteStart then
						dteStartDate = i_dteStart
					Else
						dteStartDate = rstLeavePeriodsInYear("datStartDate")
					End If

					If rstLeavePeriodsInYear("datEndDate") > i_dteEnd then
						dteEndDate = i_dteEnd
					Else
						dteEndDate = rstLeavePeriodsInYear("datEndDate")
					End If		
																			
					dblDays = DateDiff("y", dteStartDate, dteEndDate) + 1				
											
					'*** If the start date is in the requested period, and is PM and not a weekend day or public holiday, take off half a day.
					If rstLeavePeriodsInYear("datStartDate") = dteStartDate and rstLeavePeriodsInYear("strStartTime") = "PM" then
							if (not glbPublicHolidays.Contains(rstLeavePeriodsInYear("datStartDate"), False)) _
								and (not mIsWeekendDay(rstLeavePeriodsInYear("datStartDate"))) then
									dblDays = dblDays - 0.5
							end if
					end if

					If rstLeavePeriodsInYear("datEndDate") = dteEndDate and rstLeavePeriodsInYear("strEndTime") = "AM" then				
							if (not glbPublicHolidays.Contains(rstLeavePeriodsInYear("datEndDate"), False)) _ 
								and (not mIsWeekendDay(rstLeavePeriodsInYear("datEndDate"))) then
									dblDays = dblDays - 0.5
							end if
					end if
										
					if blnBalanceAffecting  and not rstLeavePeriodsInYear("blnIsOtherLeave") then
						
						if (Not blnELP and IsNull(rstLeavePeriodsInYear("lngELPID"))) or blnELP then
							dblLeaveDaysInPeriod = dblLeaveDaysInPeriod + dblDays - glbPublicHolidays.CountInPeriod(dteStartDate, dteEndDate,False) - mCountWeekendDays(dteStartDate, dteEndDate)
						end if
																	
					end if

					if Not blnBalanceAffecting And (blnAffectOnlyLegalAllowance And StopsLegalAdjAccrual) then
						
						dblLeaveDaysInPeriod = dblLeaveDaysInPeriod + dblDays - glbPublicHolidays.CountInPeriod(dteStartDate, dteEndDate,False) - mCountWeekendDays(dteStartDate, dteEndDate)
						
					end if

				end if
												
				rstLeavePeriodsInYear.Movenext
							
			wend
			
			LeaveDaysInPeriod = dblLeaveDaysInPeriod
			
		End Function
		
		
		Private Function OtherLeaveDaysInPeriod(ByVal i_dteStart, ByVal i_dteEnd)
		
			Dim dteStartDate
			Dim dteEndDate
			Dim dblDays
			
			Dim dblLeaveDaysInPeriod
			
			dblLeaveDaysInPeriod = 0
			
			if rstLeavePeriodsInYear.EOF then
				if not rstLeavePeriodsInYear.BOF then
					rstLeavePeriodsInYear.MoveFirst	
				end if
			else
				if not rstLeavePeriodsInYear.BOF then
					rstLeavePeriodsInYear.MoveFirst	
				end if
			end if
						
			while not rstLeavePeriodsInYear.EOF
							
				If rstLeavePeriodsInYear("datStartDate") > i_dteEnd or rstLeavePeriodsInYear("datEndDate") < i_dteStart then
		
					dblDays = dblDays + 0
		
				elseif rstLeavePeriodsInYear("blnIsOtherLeave") = 0 then
				
					dblDays = dblDays + 0
					
				else
				
					If rstLeavePeriodsInYear("datStartDate") < i_dteStart then
						dteStartDate = i_dteStart
					Else
						dteStartDate = rstLeavePeriodsInYear("datStartDate")
					End If

					If rstLeavePeriodsInYear("datEndDate") > i_dteEnd then
						dteEndDate = i_dteEnd
					Else
						dteEndDate = rstLeavePeriodsInYear("datEndDate")
					End If		
																			
					dblDays = DateDiff("y", dteStartDate, dteEndDate) + 1				
											
					'*** If the start date is in the requested period, and is PM and not a weekend day or public holiday, take off half a day.
					If rstLeavePeriodsInYear("datStartDate") = dteStartDate and rstLeavePeriodsInYear("strStartTime") = "PM" then
							if (not glbPublicHolidays.Contains(rstLeavePeriodsInYear("datStartDate"), False)) _
								and (not mIsWeekendDay(rstLeavePeriodsInYear("datStartDate"))) then
									dblDays = dblDays - 0.5
							end if
					end if

					If rstLeavePeriodsInYear("datEndDate") = dteEndDate and rstLeavePeriodsInYear("strEndTime") = "AM" then				
							if (not glbPublicHolidays.Contains(rstLeavePeriodsInYear("datEndDate"), False)) _ 
								and (not mIsWeekendDay(rstLeavePeriodsInYear("datEndDate"))) then
									dblDays = dblDays - 0.5
							end if
					end if
										
						
					dblLeaveDaysInPeriod = dblLeaveDaysInPeriod + dblDays - glbPublicHolidays.CountInPeriod(dteStartDate, dteEndDate,False) - mCountWeekendDays(dteStartDate, dteEndDate)

				end if
												
				rstLeavePeriodsInYear.Movenext
							
			wend
			
			OtherLeaveDaysInPeriod = dblLeaveDaysInPeriod
			
		End Function
		
		
		
		
		'** Public Methods
		
		Public Sub Initialise(ByRef i_rsEmployeeBaseDetails)
			
			strWWID = i_rsEmployeeBaseDetails("WWID")
			strStatus = i_rsEmployeeBaseDetails("EmployeeStatusCd")
			strExempt = i_rsEmployeeBaseDetails("FLSACd")
			strName = i_rsEmployeeBaseDetails("FullName")
			strManagerName = i_rsEmployeeBaseDetails("NextLevelNm")
			strSite = i_rsEmployeeBaseDetails("WorkLocationSiteCd")
            dteOriginalStartDt = CDate(i_rsEmployeeBaseDetails("OriginalStartDt"))
			dteLastDateOfHire = CDate(i_rsEmployeeBaseDetails("StartDt"))
            
			blnBlueBadge = i_rsEmployeeBaseDetails("blnBlueBadge")
			dteDOB = i_rsEmployeeBaseDetails("datDOB")
			dteEndDate = i_rsEmployeeBaseDetails("endDate")
			strOrgUnit = i_rsEmployeeBaseDetails("DepartmentNm")
			
			intNumVacationTaken = 0 ' Added by [MOF 11/08]
		End Sub
		
		Public Sub LoadValues(ByVal i_dteStartDate, ByVal i_dteEndDate)
		
			Dim intTotalLeaveToDate, intLeaveBookedToTake
			Dim cmGetEmployeeLeaveToDate, rsEmployeeLeaveToDate
			Dim cmGetEmployeeLeaveForYear
			Dim cmGetEmployeeCarryOver, rsEmployeeCarryOver
			Dim cmGetELPData, rsELPData
			
			Dim lngEEStartDateDay
			Dim lngDaysInEEsFirstMonth
			Dim lngEEWholeMonthsInYear
			Dim lngMonthlyEntitlement

			Dim lngQualLOS
			Dim lngQualAge
			
			Dim intDaysUntilEndOfYear
			Dim intWorkingDaysUntilEndOfYear
			Dim intLeaveBookedUntilEndOfYear
			
			Dim dblFraction
			
			intYear = DatePart("yyyy", i_dteStartDate)
			dteStartOfYear = mFirstDayOfYear(intYear)
			dteEndOfYear = mLastDayOfYear(intYear)
			dteSelectedStartDate = i_dteStartDate
			dteSelectedEndDate = i_dteEndDate
			
			
			
			lngDaysInYear = mGetDaysInYear(intYear)
			lngWeekendDaysInYear = mCountWeekendDays(dteStartOfYear, dteEndOfYear)
			lngPublicHolidaysInYear = glbPublicHolidays.CountInPeriod(dteStartOfYear, dteEndOfYear, False)
			
			
			lngWorkingDaysInYear =  lngDaysInYear - lngPublicHolidaysInYear - lngWeekendDaysInYear
			dteStartOfNextMonth = DateAdd("d", 1, dteSelectedEndDate)
			
			intDaysUntilEndOfYear = DateDiff("d", dteStartOfNextMonth, dteEndofYear) + 1
			
			
			'*********************************************
			'**				
			'** retrieve the employees leave for the year
			'**
			'*********************************************
			Set cmGetEmployeeLeaveForYear = Server.CreateObject("ADODB.Command")
			Set cmGetEmployeeLeaveForYear.ActiveConnection =  m_cnDB
			cmGetEmployeeLeaveForYear.CommandType = 4
			cmGetEmployeeLeaveForYear.CommandText = "dbo.pr_evc_employee_leave_for_year"
			cmGetEmployeeLeaveForYear.Parameters.Append cmGetEmployeeLeaveForYear.CreateParameter("@vWWID", adChar, adParamInput, 8, CStr(strWWID))
			cmGetEmployeeLeaveForYear.Parameters.Append cmGetEmployeeLeaveForYear.CreateParameter("@vYear", adSmallInt, adParamInput, ,CInt(intYear))
			Set rstLeavePeriodsInYear = cmGetEmployeeLeaveForYear.Execute
			
			
			'*********************************************
			'**				
			'** determine base entitlement
			'**
			'*********************************************
        ' return the basic entitlement for the current user this year      
        ' calculated if user has started or will finish this year (not working all the year) 
		 If blnBlueBadge then
				dim LeaveTypeRules
				set LeaveTypeRules = new cObjLeaveType
				LeaveTypeRules.Name = CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION

           Dim loclngEEStartDateDay
            Dim loclngEEEndDateDay ' last day of work
            Dim loclngFirstMonthEntitlement 
            Dim loclngLastMonthEntitlement ' last day of work
            
            Dim loclngDaysInEEsFirstMonth
            Dim loclngDaysInEEsLastMonth
            Dim loclngEEWholeMonthsInYear
            
            Dim loclngMonthlyEntitlement

            
            	if (IsDate(dteEndDate) and  dteEndDate< dteStartOfYear) then
					dblBasicEntitlement = 0
                elseif (dteLastDateOfHire >= dteStartOfYear) or (IsDate(dteEndDate) and  dteEndDate< dteEndOfYear) then
                'specific basic entitlement
					loclngMonthlyEntitlement = (int((LeaveTypeRules.EntitlementAmount / 12)*100))/100
					If dteLastDateOfHire >= dteStartOfYear then
						' starting working this year
						loclngEEStartDateDay = dteLastDateOfHire
						loclngDaysInEEsFirstMonth = mGetDaysInMonthForDate(dteLastDateOfHire)
						loclngFirstMonthEntitlement = (loclngDaysInEEsFirstMonth - datepart("d",dteLastDateOfHire)) + 1 'nb of days
                    	loclngFirstMonthEntitlement = ((loclngFirstMonthEntitlement * loclngMonthlyEntitlement) / loclngDaysInEEsFirstMonth)
					else
						loclngFirstMonthEntitlement = loclngMonthlyEntitlement
						loclngEEStartDateDay = dteStartOfYear
					end if
			if (IsDate(dteEndDate) and  dteEndDate< dteEndOfYear) then
						' finishing working this year
						loclngEEEndDateDay = dteEndDate
						loclngDaysInEEsLastMonth = mGetDaysInMonthForDate(dteEndDate)
						loclngLastMonthEntitlement = (datepart("d",dteEndDate) * loclngMonthlyEntitlement) / loclngDaysInEEsLastMonth
			
					else 
						loclngLastMonthEntitlement = loclngMonthlyEntitlement
						loclngEEEndDateDay = dteEndOfYear
					end if
			
					' special case if the same month for both
                    if ((DatePart("m", loclngEEEndDateDay)= DatePart("m", loclngEEStartDateDay)) and (DatePart("yyyy", loclngEEEndDateDay)= DatePart("yyyy", loclngEEStartDateDay))) then
						loclngFirstMonthEntitlement = (loclngEEEndDateDay - loclngEEStartDateDay) + 1 'nb of days
						dblBasicEntitlement = (loclngFirstMonthEntitlement * loclngMonthlyEntitlement)/mGetDaysInMonthForDate(loclngEEStartDateDay)
                    else	                    
						loclngEEWholeMonthsInYear = ((DatePart("m", loclngEEEndDateDay)-1)) - (DatePart("m", loclngEEStartDateDay)) '[MFILLAST 08-2006] chge 12 in datepart("m",EE.EndsDate)
						' add full month entitlement
						dblBasicEntitlement = (loclngLastMonthEntitlement + loclngFirstMonthEntitlement + (loclngEEWholeMonthsInYear * loclngMonthlyEntitlement))                   
					end if
            			dblBasicEntitlement = mRoundToNextHalf(dblBasicEntitlement)
			    else ' working all the year
                   dblBasicEntitlement = LeaveTypeRules.EntitlementAmount
                end if
            else
            				dblBasicEntitlement = 0
			End If
    				
				
			'************************************
			'**				
			'** determine carry over
			'**
			'************************************
			set cmGetEmployeeCarryOver = Server.CreateObject("ADODB.Command")
			set cmGetEmployeeCarryOver.ActiveConnection =  m_cnDB
			cmGetEmployeeCarryOver.CommandType = 4
			cmGetEmployeeCarryOver.CommandText = "dbo.pr_evc_employee_carry_over_for_year"
			cmGetEmployeeCarryOver.Parameters.Append cmGetEmployeeCarryOver.CreateParameter("@vWWID", adChar, adParamInput, 8, CStr(strWWID))
			cmGetEmployeeCarryOver.Parameters.Append cmGetEmployeeCarryOver.CreateParameter("@vYear", adSmallInt, adParamInput, , CInt(intYear))
			set rsEmployeeCarryOver = cmGetEmployeeCarryOver.Execute
			
			if rsEmployeeCarryOver.EOF and rsEmployeeCarryOver.BOF then
				dblCarryOver = 0 
			elseif IsNull(rsEmployeeCarryOver("NumCarried")) then
				dblCarryOver = 0
			else
				dblCarryOver = rsEmployeeCarryOver("NumCarried")
			end if
			
			set cmGetEmployeeCarryOver = nothing
			set rsEmployeeCarryOver = nothing
			

						
			'************************************
			'**				
			'** determine total leave for year
			'**
			'************************************
			dblTotalLeaveBooked = LeaveDaysInPeriod(dteStartOfYear, dteEndOfYear, False, True, False)

			'************************************
			'**				[MFILLAST 08-2006]
			'** determine ELP booked this year 
			'** cacul is : total leaves including ELP - total leaves without ELP
			'************************************
			dblELPUsed = LeaveDaysInPeriod(dteStartOfYear, dteEndOfYear, False, True, true) - dblTotalLeaveBooked
	
			'************************************
			'**				
			'** determine ELP banked
			'**
			'************************************
			'[MFILLAST 08-2006] : bug here when several entries -> SOLVED using a loop
			set cmGetELPData = Server.CreateObject("ADODB.Command")
			set cmGetELPData.ActiveConnection =  m_cnDB
			cmGetELPData.CommandType = 4
			cmGetELPData.CommandText = "dbo.pr_evc_elp_data_for_employee"
			cmGetELPData.Parameters.Append cmGetELPData.CreateParameter("@vWWID", adChar, adParamInput, 8, CStr(strWWID))
			set rsELPData = cmGetELPData.Execute
			
			dblELPBanked = 0
			if not rsELPData.EOF then 'now looking for line with good information inside
				'[MFILLAST 08-2006] : added procedures to take care of relief
				dim cmGetELPRelief,rsELPRelief
				set cmGetELPRelief = Server.CreateObject("ADODB.Command")
				set cmGetELPRelief.ActiveConnection =  m_cnDB
				cmGetELPRelief.CommandType = 4
				
				dim locBlnFound 
				locBlnFound = false
				Do Until (rsELPData.EOF or locBlnFound) 'stop when value found
				
					if DatePart("yyyy", rsELPData("datActivated")) = intYear then 'year of activation
						dblELPBanked = rsELPData("dblInitialDaysBanked")
						locBlnFound = true
					else
						Dim intYearsSinceStart
						Dim intDaysToBank
						Dim intYearsToRun
						'check if there is a relief this year : no days banked but elp found
						
						cmGetELPRelief.CommandText = "dbo.get_elprelief_year"
						Do While (cmGetELPRelief.Parameters.Count > 0)
							cmGetELPRelief.Parameters.Delete 0
						Loop
						cmGetELPRelief.Parameters.Append cmGetELPRelief.CreateParameter("@lngYear", adSmallInt, adParamInput, , intYear)
						cmGetELPRelief.Parameters.Append cmGetELPRelief.CreateParameter("@lngELPID", adInteger, adParamInput, , rsELPData("lngID"))
						set rsELPRelief = cmGetELPRelief.Execute
						'response.Write "isrelief:"& rsELPRelief("isRelief")
						'response.Write "current year : "&intYear
						'response.Write "lngELPID:"&rsELPData("lngID")
						if (rsELPRelief("isRelief")>0) then
							dblELPBanked = 0
							locBlnFound = true
						else
							cmGetELPRelief.CommandText = "dbo.get_nb_elprelief"
							Do While (cmGetELPRelief.Parameters.Count > 0)
								cmGetELPRelief.Parameters.Delete 0
							Loop
							cmGetELPRelief.Parameters.Append cmGetELPRelief.CreateParameter("@lngELPID", adInteger, adParamInput, , rsELPData("lngID"))
							set rsELPRelief = cmGetELPRelief.Execute
					'	response.Write "nbRelief:"& rsELPRelief("nbRelief")
						
							intYearsSinceStart = DateDiff("yyyy", rsELPData("datActivated"), dteEndOfYear)
							intDaysToBank = (rsELPData("dblTargetDays") - rsELPData("dblInitialDaysBanked"))
							intYearsToRun = intDaysToBank/5 ' 5 days are banked each years afetr the first
							intYearsToRun = intYearsToRun + rsELPRelief("nbRelief") ' each relief add one year
							if intYearsSinceStart > intYearsToRun then 'old ELP, useless
								dblELPBanked = 0
							else
								dblELPBanked = 5
								locBlnFound = true
							end if						
						end if
						

					end if
					rsELPData.MoveNext
				Loop
				set cmGetELPRelief = Nothing
				set rsELPRelief = nothing 
			end if
			
			set cmGetELPData = Nothing
			set rsELPData = nothing
			

			'*************************************
			'**
			'** Determine Leave Balance at EOY
			'**
			'*************************************			
			intLeaveBalanceAtEndOfYear = dblBasicEntitlement + dblCarryOver - dblTotalLeaveBooked - dblELPBanked '[MFILLAST 08-2006] delete + dblTotalRestdays 
			
			'*************************************
			'**
			'** Determine Number of Days Worked
			'**
			'*************************************
			if dteLastDateOfHire > dteStartOfYear then
				intNumberOfDaysWorked = DateDiff("d", dteLastDateOfHire, i_dteEndDate) + 1
				intNumberOfDaysWorked = intNumberOfDaysWorked - mCountWeekendDays(dteLastDateOfHire, i_dteEndDate)
				intNumberOfDaysWorked = intNumberOfDaysWorked - glbPublicHolidays.CountInPeriod(dteLastDateOfHire, i_dteEndDate, False)
				intNumberOfDaysWorked = intNumberOfDaysWorked - LeaveDaysInPeriod(dteLastDateOfHire, i_dteEndDate, False, True, True)
				intNumberOfDaysWorked = intNumberOfDaysWorked - OtherLeaveDaysInPeriod(dteStartOfYear, i_dteEndDate)
			else 
				intNumberOfDaysWorked = DateDiff("d", dteStartOfYear, i_dteEndDate) + 1
				intNumberOfDaysWorked = intNumberOfDaysWorked - mCountWeekendDays(dteStartOfYear, i_dteEndDate)
				intNumberOfDaysWorked = intNumberOfDaysWorked - glbPublicHolidays.CountInPeriod(dteStartOfYear, i_dteEndDate, False)
				intNumberOfDaysWorked = intNumberOfDaysWorked - LeaveDaysInPeriod(dteStartOfYear, i_dteEndDate, False, True, True)
				intNumberOfDaysWorked = intNumberOfDaysWorked - OtherLeaveDaysInPeriod(dteStartOfYear, i_dteEndDate)
			end if
			
			
			
			
			
			
			'********************************************
			'**
			'** Determine leave booked until end of year
			'**
			'********************************************
			intLeaveBookedUntilEndOfYear =  LeaveDaysInPeriod(dteStartOfNextMonth, dteEndOfYear, False, True, True)
			intLeaveBookedUntilEndOfYear = intLeaveBookedUntilEndOfYear + OtherLeaveDaysInPeriod(dteStartOfNextMonth, dteEndOfYear)
			
			
			'*************************************************
			'**
			'** Determine total vacation taken
			'**
			'*************************************************		
			' intNumVacationTaken = 
			' TO DO: FIX ABOVE
			
			' Changed "Taken" to "Approved", calculating "Taken" days using the datConfirmed column [MOF 11/25/08] 
			'*************************************************
			'**
			'** Determine total vacation approved 
			'**
			'*************************************************		
			intNumVacationApproved = LeaveDaysInPeriod(dteStartOfYear, dteSelectedEndDate, False, True, false)'[MFILLAST 08-2006] do not count ELP : True)
			
			'*************************************************
			'**
			'** Determine other leave  vacation taken
			'**
			'*************************************************						
			intNumOtherLeaveTaken = OtherLeaveDaysInPeriod(dteStartOfYear, dteSelectedEndDate)
			
			
			'*************************************************
			'**
			'** Determine legal carry over
			'**
			'*************************************************						
		'[MFILLAST 08-2006] : carry over = balance
			intCarryOverEOY = intLeaveBalanceAtEndOfYear
			
			set rstLeavePeriodsInYear = nothing
			
		End Sub
		
	End Class
	
'[MFILLAST 08-2006]	Dim glbPublicHolidays

	Dim m_cnDB
	Dim m_cmEmployeeBaseData, m_rsEmployeeBaseData
	Dim m_objEmployee

	Dim dteSelectedStartDate, dteSelectedEndDate
	Dim strStatus, strSite, strExempt, strEmployee, strManager
	
	Dim blnEmpSearch
	Dim blnMgrSearch
	Dim blnShow
	
	Dim strMonth
	Dim strYear
	
	'create a new connection because need to modify cursor location(error otherwise)
	set m_cnDB = Server.CreateObject("ADODB.Connection")
	'm_cnDB.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=evac_web;Data Source=irea1sqlp200;Initial Catalog=dbn_evacation;Password=legislation;" 'comment removed by slenno2 27 july
	m_cnDB.ConnectionString = CONST_ADO_EVACATION_CONNECTION_STRING	
	m_cnDB.CursorLocation = adUseClient
	m_cnDB.Open
	
	strMonth = Request.QueryString("datMonth")
	strYear = Request.QueryString("datYear")
	
	'dteSelectedStartDate = Request.QueryString("datStartDate")
	dteSelectedEndDate = mLastDayOfMonth(strMonth, strYear)
	dteSelectedStartDate = dteSelectedEndDate 
	
	strStatus = Request.QueryString("status")
	strExempt = Request.QueryString("exempt")
	strSite = Request.QueryString("site")
	strEmployee = Request.QueryString("employee")
	'strEmployee = "10664024"
	strManager = Request.QueryString("manager")
 

	
	blnEmpSearch = False
	blnMgrSearch = False
	
	if Not IsNull(strEmployee) then
		if trim(strEmployee) > "" then
			blnEmpSearch = True
		end if
	end if
	
	if Not IsNull(strManager) then
		if trim(strManager) > "" then
			blnMgrSearch = True
		end if
	end if
	
	' no manager and employee at the same time : error page
	if blnEmpSearch and blnMgrSearch then
	
		response.write "<HTML>"
		response.write "<HEAD>"
			response.write "<TITLE>"
				response.write "e-Vacation - "
				response.write "HR Report"
			response.write "</TITLE>"
			response.write "<link rel=STYLESHEET type=""text/css"" href=""" & CONST_APPLICATION_PATH & "common/css/evacation.css"" TITLE=""style"">"
			response.write "</HEAD>"
			response.write "<BODY>"
			
			response.write "<center>"
				response.write "<br>"
				response.write "<font class=error>"
					response.write "<b>"
						response.write "You have entered both a manager and an employee. Please select only one.<br><br>Please close this window and try again.<br>"
					response.write "</b><br>"
				response.write "</font>"
				response.write "<br>"
			response.write "</center>"
			
		response.write "<table width=700 border=0 cellspacing=0 cellpadding=3 align=center class=nav>"
			response.write "<tr>"
				response.write "<td style='text-align:center'>"
					response.write "&copy;2001 Intel Corporation - Developed by GE HRIS/TIS"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		response.write "</body>"
		response.write "</html>"
		
		Response.End
		
	end if
	
	
	
	
	set m_cmEmployeeBaseData = Server.CreateObject("ADODB.Command")
	set m_cmEmployeeBaseData.ActiveConnection = m_cnDB
	m_cmEmployeeBaseData.CommandType = adCmdStoredProc
	if blnEmpSearch then
		m_cmEmployeeBaseData.CommandText = "dbo.pr_evc_hr_report_employee_data"
		m_cmEmployeeBaseData.Parameters.Append m_cmEmployeeBaseData.CreateParameter("@vWWID", adChar, adParamInput, 8, CStr(strEmployee))
	elseif blnMgrSearch then
		m_cmEmployeeBaseData.CommandText = "dbo.pr_evc_hr_report_manager_reports"
		m_cmEmployeeBaseData.Parameters.Append m_cmEmployeeBaseData.CreateParameter("@vWWID", adChar, adParamInput, 8, CStr(strManager))
	else
		m_cmEmployeeBaseData.CommandText = "dbo.pr_evc_hr_report_base_data"
	end if
	set m_rsEmployeeBaseData = m_cmEmployeeBaseData.Execute
'[MFILLAST 08-2006]	Set glbPublicHolidays = new cColPublicHolidays
	

	


'[MFILLAST 08-2006] debug in non excel mode
'		response.write "<HTML>"
'		response.write "<HEAD>"
'			response.write "<TITLE>"
'				response.write "e-Vacation - "
''				response.write "HR Report"
'		response.write "</TITLE>"
'			response.write "<link rel=STYLESHEET type=""text/css"" href=""" & CONST_APPLICATION_PATH & "common/css/evacation.css"" TITLE=""style"">"
'			response.write "</HEAD>"
'			response.write "<BODY>"
			'/[MFILLAST 08-2006]
	
	'***** Set content-type to Excel (for display in browser as Excel workbook). *******
	Response.ContentType = "application/vnd.ms-excel" 
	
	'***** SET THIS HTML DOCUMENT TO BE RECOGNISED BY THE BROWSER AS AN EXCEL SPREADSHEET ******
	response.write "<HTML xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
		
		'***** SET UP PAGE ****
		response.write "<HEAD>"
			response.write "<style>"
		  	response.write "<!--table"
				  	response.write "@page"
			     	response.write "{mso-header-data:""&Ce-Vacation - HR Report\000A"
	
			     	response.write "Report Period: " & MonthName(DatePart("m", dteSelectedEndDate), true) & "/" & DatePart("yyyy", dteSelectedEndDate) & "\000A"
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
				      	response.write "<x:Name>HR Report</x:Name>"
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
	
		'** Display Header **
		response.write "<TABLE>"
		response.write "<TR>"
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "WWID" 'CA
			response.write "</TD>"
    
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "Employee Status" 'CA
			response.write "</TD>"
    
'			response.write "<TD BGColor=""#c0c0c0"">"
'				Response.Write "Exempt Status"
'			response.write "</TD>"
    
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "Employee Name"
			response.write "</TD>"
    
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "Manager Name"
			response.write "</TD>"
    
    
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "Org Unit"
			response.write "</TD>"
    
'			response.write "<TD BGColor=""#c0c0c0"">"
'				Response.Write "Site"
'			response.write "</TD>"
			
			'[MFILLAST 08-2006]
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "Entitlement" 
			response.write "</TD>"
    
			'[MFILLAST 08-2006]
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "Carryover/Exceptions" 
			response.write "</TD>"
    
			'[MFILLAST 08-2006]
			response.write "<TD BGColor=""#c0c0c0""><b>"
				Response.Write "Initial Balance" 
			response.write "</b></TD>"
    
			response.write "<TD BGColor=""#c0c0c0"">"
				'response.write .DaysBooked 
				Response.Write "Vacation Taken/Approved"
			response.write "</TD>"
			
		'	response.write "<TD BGColor=""#c0c0c0"">"
		'		Response.Write "Vacation Approved"
		'	response.write "</TD>"
			
		'	response.write "<TD BGColor=""#c0c0c0"">"
		'		Response.Write "Vacation Taken"
		'	response.write "</TD>"
			
		'	response.write "<TD BGColor=""#c0c0c0"">"
		'		Response.Write "Comp. Days Taken"
		'	response.write "</TD>"
			
		'	response.write "<TD BGColor=""#c0c0c0"">"
		'		Response.Write "<b>Total Days Taken</b>"
		'	response.write "</TD>"
        
			'[MFILLAST 08-2006]
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "Deposit to ELP" 
			response.write "</TD>"
    
			'[MFILLAST 08-2006]
			'response.write "<TD BGColor=""#c0c0c0""><b>"
			'	Response.Write "Consumption" 
			'response.write "</b></TD>"
    
			response.write "<TD BGColor=""#c0c0c0""><b>"
				Response.Write "Leave Balance at EOY" 
			response.write "</b></TD>"
    
'			response.write "<TD BGColor=""#c0c0c0"">"
'				Response.Write "Number of Days Worked"
'			response.write "</TD>"
    
    
'			response.write "<TD BGColor=""#c0c0c0"">"
'				Response.Write "Picture EOY"
'			response.write "</TD>"
    
					
			response.write "<TD BGColor=""#c0c0c0"">"
				Response.Write "ELP Taken/Approved"
			response.write "</TD>"					
					
'			response.write "<TD BGColor=""#c0c0c0"">"
'				Response.Write "Carry Over"
'			response.write "</TD>"	
											
		response.write "</TR>"
		'** Finish Display Header **
		
	while not m_rsEmployeeBaseData.EOF

			if trim(LCase(strStatus)) <> "all" then
			
				if trim(LCase(strStatus)) = "active" then
                    if m_rsEmployeeBaseData("EmployeeStatusCd") = "A" then
					    blnShow = True
                    else
						blnShow = False
					end if				
				end if
				
				if trim(LCase(strStatus)) = "terminated" then
					blnShow = False
                    if m_rsEmployeeBaseData("EmployeeStatusCd") = "R" or m_rsEmployeeBaseData("EmployeeStatusCd") = "T" then
					    blnShow = True
                    else
						blnShow = False
					end if
                end if
    		else
                 blnShow = True
			end if
				
		if blnShow then
		
			set m_objEmployee = new clsEmployeeHRData
			m_objEmployee.Initialise m_rsEmployeeBaseData
                m_objEmployee.LoadValues CDate(dteSelectedStartDate), CDate(dteSelectedEndDate)
            
			Response.Write "<TR><TD>"
			Response.Write m_objEmployee.WWID 
			Response.Write "</TD><TD>"
			Response.Write m_objEmployee.Status
			Response.Write "</TD><TD>"
'			Response.Write m_objEmployee.Exempt 
'			Response.Write "</TD><TD>"
			Response.Write m_objEmployee.Name 
			Response.Write "</TD><TD>"
			Response.Write m_objEmployee.Manager 
			Response.Write "</TD><TD>"
			Response.Write m_objEmployee.OrgUnit 
'			Response.Write "</TD><TD>"
'			Response.Write m_objEmployee.Site 
			Response.Write "</TD><TD>"
			Response.Write m_objEmployee.Entitlement
			Response.Write "</TD><TD>"
			Response.Write m_objEmployee.PreviousCarryOver
			Response.Write "</TD><TD>"
			Response.Write m_objEmployee.Entitlement + m_objEmployee.PreviousCarryOver
			Response.Write "</TD><TD>"
			'Response.Write "X"
			Response.Write m_objEmployee.Entitlement + m_objEmployee.PreviousCarryOver - m_objEmployee.BalanceAtEndOfYear		
			Response.Write "</TD><TD>"
			'Response.Write m_objEmployee.VacationApproved 		
			'Response.Write "</TD><TD>"
			'Response.Write m_objEmployee.VacationTaken 		
			'Response.Write "</TD><TD>"
			'Response.Write "-"
			'Response.Write m_objEmployee.VacationTaken 		
			'Response.Write "</TD><TD>"
			'Response.Write "-"
			'Response.Write "</TD><TD>"
			Response.Write m_objEmployee.DepositELP
			Response.Write "</TD><TD>"
			'Response.Write m_objEmployee.VacationTaken + m_objEmployee.DepositELP
			'Response.Write "</TD><TD>"
			Response.Write m_objEmployee.BalanceAtEndOfYear
			Response.Write "</TD><TD>"
		    Response.Write m_objEmployee.ELPUsed 
			Response.Write "</TD><TD>"
		
			set m_objEmployee = nothing
		
		end if
    
		m_rsEmployeeBaseData.MoveNext
		
		'Response.Flush
				
	wend
	
	Response.Write "</TABLE></BODY></HTML>"
	
	set m_rsEmployeeBaseData = nothing
	
%>
