<%
    '***** FUNCTIONS *****
    '
    'mCalculatedIsEELeaveTracked    (ByRef locobjUser)
    'mCleanLineBreaks               (ByVal locstrValue)
    'mCountWeekendDays              (ByVal locdatStartDate, ByVal locdatEndDate)
    'mCountWholeYears               (ByVal locdatFirst, ByVal locdatSecond)
    'mFirstDayOfMonth               (ByVal loclngMonth,ByVal loclngYear)
    'mFirstDayOfYear                (ByVal loclngYear)
    'mFormatDate                    (ByVal locdatDate, ByVal strOption)
    'mGetDaysInMonthForDate         (ByVal locdatDate)
    'mGetDaysInYear                 (ByVal loclngYear)
    'mGetSafeLongInteger            (ByVal locvarValue, ByVal varReturnInvalid)
    'mHTMLEncode                    (ByVal locstrValue)
    'mIf                            (ByVal locblnValue, ByVal locvarReturnTrue, ByVal locvarReturnFalse)
    'mInitialiseCurrentUser         ()
    'mInitialiseEEtoView            ()
    'mIsEmptyString                 (ByVal locstrString)
    'mIsFormattedDate               (ByVal locvarDate)
    'mIsValidWWID                   (ByVal locstrString)
    'mIsWeekendDay                  (ByVal locdatDate)
    'mLastDayOfYear                 (ByVal loclngYear)
    'mLastDayOfMonth                (ByVal loclngMonth,ByVal loclngYear)
    'mLoadCurrentUser               ()
    'mRoundToNextHalf               (ByVal loclngValue)
    'mSendEmail                     (ByVal locstrFrom, ByVal locstrTo, ByVal locstrSubject, ByVal locstrBody, ByVal locblnHTML)
	'updateEmployees 				(ByRef locobjEEtoView)
    

    '*** CALCULATE IS EE LEAVE TRACKED ***
    function mCalculateIsEELeaveTracked(ByRef locobjUser)
        With locobjUser

            If .IsBlueBadge _
                AND .CompanyCd = "508" _
                AND .ActiveThisYear then
                mCalculateIsEELeaveTracked = True
            Else
                mCalculateIsEELeaveTracked = False
            End If
        End With
    end function
    
    
    '*** CLEAN LINE BREAKS ***
    function mCleanLineBreaks(ByVal locstrValue)
        if mIsEmptyString(locstrValue) then
            locstrValue = ""
        else
            locstrValue = replace(locstrValue,chr(10) & chr(13), chr(13))
            locstrValue = replace(locstrValue,chr(13) & chr(10), chr(13))
            locstrValue = replace(locstrValue,chr(10), chr(13))
            locstrValue = replace(locstrValue,"<BR>","<br>")
            locstrValue = replace(locstrValue,"<Br>","<br>")
            locstrValue = replace(locstrValue,"<bR>","<br>")
            locstrValue = replace(locstrValue,"<br>" & chr(13), chr(13))
            locstrValue = replace(locstrValue,"<br>", chr(13))
            'Remove any line-breaks followed by spaces
            locstrValue = replace(locstrValue,chr(13) & " ", chr(13))
            'Remove any triple line-breaks
            while instr(locstrValue,chr(13) & chr(13) & chr(13)) > 0
                locstrValue = replace(locstrValue,chr(13) & chr(13) & chr(13), chr(13) & chr(13))
            wend
        end if
        mCleanLineBreaks = locstrValue
    end function
	

    '*** COUNT WEEKEND DAYS ***
    function mCountWeekendDays(ByVal locdatStartDate, ByVal locdatEndDate)
        '*** WRITTEN BY CHRISTINE ARMSTRONG - 1 April 2001 ****
        Dim locDatInputStartDate
        Dim locDatInputEndDate
        Dim locDayOfInputStartDate
        Dim locDayOfInputEndDate
        Dim locIntNumberOfSundays
        Dim locIntNumberOfWeekendDays
        
        locIntNumberOfWeekendDays = 0

        If (not IsDate(locdatStartDate)) or (not IsDate(locdatEndDate)) then
            mCountWeekendDays = 0
            Exit Function
        End If
        
        'If interval is Week ("ww"), however, the DateDiff function returns
        'the number of calendar weeks between the two dates.
        'It counts the number of Sundays between date1 and date2.
        'DateDiff counts date2 if it falls on a Sunday;
        'but it doesn't count date1, even if it does fall on a Sunday.
        'The firstdayofweek argument affects calculations that use the
        '"w" and "ww" interval symbols.
        
        locDatInputStartDate = CDate(locdatStartDate)
        locDatInputEndDate = CDate(locdatEndDate)
        
        locDayOfInputStartDate = DatePart("w", locDatInputStartDate, vbSunday, vbFirstJan1)
        locDayOfInputEndDate = DatePart("w", locDatInputEndDate, vbSunday, vbFirstJan1)
        
        locIntNumberOfSundays = DateDiff("ww", locDatInputStartDate, locDatInputEndDate, vbSunday, vbFirstJan1)
        locIntNumberOfWeekendDays = locIntNumberOfSundays * 2
        
        If locDayOfInputStartDate = vbSunday Then
            locIntNumberOfWeekendDays = locIntNumberOfWeekendDays + 1
        End If
        
        If locDayOfInputEndDate = vbSaturday Then
            locIntNumberOfWeekendDays = locIntNumberOfWeekendDays + 1
        End If
        
        mCountWeekendDays = locIntNumberOfWeekendDays
            
    end function


    '*** COUNT WHOLE YEARS ***
    function mCountWholeYears(ByVal locdatFirst, ByVal locdatSecond)
        Dim loclngYears
        If isdate(locdatFirst) and isdate(locdatSecond) then
            loclngYears = datediff("yyyy",locdatFirst,locdatSecond)
            If DateSerial(Year(locdatSecond), Month(locdatFirst), Day(locdatFirst)) > locdatSecond Then
                loclngYears = loclngYears - 1
            End If
            if loclngYears < 0 then loclngYears = 0
            mCountWholeYears = loclngYears
        else
            mCountWholeYears = null
        end if
    end function


    '*** FIRST DAY OF MONTH ***
    function mFirstDayOfMonth(ByVal loclngMonth,ByVal loclngYear)
        If (mGetSafeLongInteger(loclngMonth,0) <> 0) and _
            (mGetSafeLongInteger(loclngYear,0) <> 0) then
            mFirstDayOfMonth = DateValue("1 " & monthname(loclngMonth) & " " & loclngYear)
        else
            mFirstDayOfMonth = ""
        End If
    end function
    

    '*** FIRST DAY OF YEAR ***
    function mFirstDayOfYear(ByVal loclngYear)
        If mGetSafeLonginteger(loclngYear,0) <> 0 then
            mFirstDayOfYear = DateValue("1 January " & loclngYear)
        Else
            mFirstDayOfYear = ""
        End If
    end function
    

    '*** FORMAT DATE ***
    function mFormatDate(ByVal locdatDate, ByVal strOption)
        if isdate(locdatDate) then
            Select Case strOption
                Case "datepicker"
                    if day(locdatDate) < 10 then
                        mFormatDate = "0"
                    end if
                    mFormatDate = mFormatDate & day(locdatDate) & "/"
                    if month(locdatDate) < 10 then
                        mFormatDate = mFormatDate & "0"
                    end if
                    mFormatDate = year(locdatDate) & "," & (month(locdatDate)-1) & "," & day(locdatDate)
                Case "short"
                    if day(locdatDate) < 10 then
                        mFormatDate = "0"
                    end if
                    mFormatDate = mFormatDate & day(locdatDate) & "/"
                    if month(locdatDate) < 10 then
                        mFormatDate = mFormatDate & "0"
                    end if
                    mFormatDate = mFormatDate & month(locdatDate) & "/" & right(cstr(year(locdatDate)),2)
                Case "medium", "medium with day", "medium with time"
                    mFormatDate = day(locdatDate)
                    mFormatDate = mFormatDate & " " & left(MonthName(DatePart("m", locdatDate)),3)
                    mFormatDate = mFormatDate & " " & Year(locdatDate)
                    
                    If strOption = "medium with day" then
                        mFormatDate = weekdayname(weekday(locdatDate),True) & ", " & mFormatDate
                    End If
                    
                    If strOption = "medium with time" then
                        mFormatDate = mFormatDate & "  " & FormatDateTime(locdatDate, vbShortTime)
                    end if
                
                Case else
                    mFormatDate = day(locdatDate) & " " & monthname(datepart("m", locdatDate)) & " " & year(locdatDate)
            End Select
        else
            mFormatDate = locdatDate
        end if
    end function

    
    '*** GET DAYS IN MONTH FOR DATE ***
    function mGetDaysInMonthForDate(ByVal locdatDate)
        Dim loclngMonth
        Dim loclngYear
        Dim locdatTempDate
        
        If not isdate(locdatDate) then
            mGetDaysInMonthForDate = 0
        Else
            loclngMonth = datepart("m",locdatDate)
            loclngYear = datepart("yyyy",locdatDate)
            locdatTempDate = datevalue("1 " & monthname(loclngMonth) & " " & loclngYear)
            
            locdatTempDate = dateadd("m",1,locdatTempDate)
            locdatTempDate = dateadd("y",-1,locdatTempDate)
            mGetDaysInMonthForDate = datepart("d",locdatTempDate)
        End If
        
    end function


    '*** GET DAYS IN YEAR ***
    function mGetDaysInYear(ByVal loclngYear)
        mGetDaysInYear = datediff("y", mFirstDayOfYear(loclngYear), mLastDayOfYear(loclngYear)) + 1
    end function
    
    
    '*** GET SAFE LONG INTEGER ***
    function mGetSafeLongInteger(ByVal locvarValue, ByVal varReturnInvalid)
        mGetSafeLongInteger = locvarValue
        if IsNumeric(locvarValue) then
            if abs(locvarValue) > 2147483647 then
                locvarValue = varReturnInvalid
            else
                locvarValue = clng(locvarValue)
            end if
        else
            locvarValue = varReturnInvalid
        end if
        mGetSafeLongInteger = locvarValue
    end function
    
    
    '*** HTML ENCODE ***
    function mHTMLEncode(ByVal locstrValue)
        if not mIsEmptyString(locstrValue) then
            locstrValue = server.htmlencode(locstrValue)
            mHTMLEncode = Replace(locstrValue,"'","&#39;")

        else
            mHTMLEncode = ""
        end if
    end function


    '*** IF ***
    function mIf(ByVal locblnValue, ByVal locvarReturnTrue, ByVal locvarReturnFalse)
        If locblnValue = True then
            mIf = locvarReturnTrue
        else
            mIf = locvarReturnFalse
        end if
    end function
    

    '*** INITIALISE CURRENT USER ***
    function mInitialiseCurrentUser()
        'CA Fixed Error handling to call user error page instead of NewUsers.asp page
        'for CONST_LOGON_ERROR_USER_SET_UP_REQUIRED, 22 May 2001
						
        lngLoadCurrentUserStatus = mLoadCurrentUser()	
		
		'dim endtime
		'endtime = timer()
		'dim benchmark
		'benchmark = endtime - starttime
		'Response.Write( benchmark ) 
		'Response.Write( " - mInitialiseCurrentUser" ) 
		'Response.Write( "<br>" ) 

        Select Case lngLoadCurrentUserStatus
            Case CONST_LOGON_SUCCESSFUL
                'No Action required
            Case CONST_LOGON_ERROR_USER_SET_UP_REQUIRED
                mCloseApplication
                response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & lngLoadCurrentUserStatus
            Case Else
                mCloseApplication
                response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & lngLoadCurrentUserStatus
        End Select
    end function


    '*** INITIALISE EE TO VIEW ***
    function mInitialiseEEtoView()
        If strRequestEEWWID = "" then strRequestEEWWID = objCurrentUser.WWID
        If strRequestEEWWID = objCurrentUser.WWID then
            Set objEEtoView = objCurrentUser
        Else
            Set objEEtoView = new cObjUser
            objEEtoView.WWID = strRequestEEWWID
        End If
    end function


    '*** IS EMPTY STRING ***
    function mIsEmptyString(ByVal locstrString)
        if typename(locstrString) = "field" then
            if len(locstrString) > 0 then
                mIsEmptyString = false
            else
                mIsEmptyString = true
            end if
        else
            if isempty(locstrString) or isnull(locstrString) or locstrString = "" then
                mIsEmptyString = true
            else
                mIsEmptyString = false
            end if
        end if
    end function


    '*** IS FORMATTED DATE ***
    function mIsFormattedDate(ByVal locvarDate)
        Dim locstrMonth
        Dim loclngLength
        
        If not IsDate(locvarDate) then
            mIsFormattedDate = False
            Exit Function
        End If
        
        loclngLength = len(locvarDate)

        If loclngLength <> 11 and loclngLength <> 10 then
            mIsFormattedDate = False
            Exit Function
        End If

        If mid(locvarDate,loclngLength-8,1) <> " " or mid(locvarDate,loclngLength-4,1) <> " " then
            mIsFormattedDate = False
            Exit Function
        End If

        locstrMonth = lcase(mid(locvarDate,loclngLength-7,3))

        If instr("jan,feb,mar,apr,may,jun,jul,aug,sep,oct,nov,dec,",locstrMonth) = 0 then
            mIsFormattedDate = False
            Exit Function
        End If
        mIsFormattedDate = True
    end function
    
    
    '*** IS VALID WWID ***
    function mIsValidWWID(ByVal locstrString)
        if len(locstrString) <> 8 then
            mIsValidWWID = false
            Exit function
        end if
        if mGetSafeLongInteger(locstrString,0) = 0 then
            mIsValidWWID = false
            Exit function
        end if
        mIsValidWWID = True
    end function
    
        
    '*** IS WEEKEND DAY ***
    function mIsWeekendDay(ByVal locdatDate)
        If (weekday(locdatDate,vbSunday) = 1) or (weekday(locdatDate,vbSunday) = 7) then
            mIsWeekendDay = True
        Else
            mIsWeekendDay = False
        End If
    end function


    '*** LAST DAY OF MONTH
    function mLastDayOfMonth(ByVal loclngMonth,ByVal loclngYear)
        If (mGetSafeLongInteger(loclngMonth,0) <> 0) and _
            (mGetSafeLongInteger(loclngYear,0) <> 0) then
            mLastDayOfMonth = DateValue(mGetDaysInMonthForDate(mFirstDayOfMonth(loclngMonth,loclngYear)) & " " & monthname(loclngMonth) & " " & loclngYear)
        else
            mFirstDayOfMonth = ""
        End If
    end function


    '*** LAST DAY OF YEAR ***
    function mLastDayOfYear(ByVal loclngYear)
        If mGetSafeLonginteger(loclngYear,0) <> 0 then
            mLastDayOfYear = DateValue("31 December " & loclngYear)
        Else
            mLastDayOfYear = ""
        End If
    end function


    '**** LOAD CURRENT USER ****
    function mLoadCurrentUser()
        Set objCurrentUser = new cObjUser		
        mLoadCurrentUser = objCurrentUser.SetToLoggedOnUser	
		
		'dim endtime
		'endtime = timer()
		'dim benchmark
		'benchmark = endtime - starttime
		'Response.Write( benchmark ) 
		'Response.Write( " - mLoadCurrentUser" ) 
		'Response.Write( "<br>" ) 
		
	end function
    

    '**** ROUND TO NEXT HALF ****
    function mRoundToNextHalf(ByVal loclngValue)
        Dim loclngFraction
        loclngFraction = loclngValue - int(loclngValue)
        If loclngFraction = 0 then
            loclngFraction = 0
        ElseIf loclngFraction <= 0.5 then
            loclngFraction = 0.5
        Else
            loclngFraction = 1
        End If
        loclngValue = int(loclngValue) + loclngFraction
        mRoundToNextHalf = loclngValue
    end function

  
   '*** SEND EMAIL USING CDO  - FOR WIN 2K3 - Added WW35 2005 Marc Doyle***
   ' Edited [MOF 12/08] Function now returns False if the mail failed to send and True otherwise
    function mSendEmail(ByVal locstrFrom, ByVal locstrTo, ByVal locstrSubject, ByVal locstrBody, ByVal locblnHTML)
        Dim objCDOMail
        Dim cdoConfig
        Dim locReturnValue
        locReturnValue = True
        'start added by [MFILLAST 08-2006]
		%>
        <!-- 
		METADATA 
		TYPE="typelib" 
		UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
		NAME="CDO for Windows 2000 Library" 
		-->  
		<%
		
        Set cdoConfig = CreateObject("CDO.Configuration")
        With cdoConfig.Fields
			.Item(cdoSendUsingMethod) = 2 'cdoSendUsingPort
			.Item(cdoSMTPServer) = CONST_MAIL_SERVER
			.update
		End With
'end added by [MFILLAST 08-2006]
        
		' [MOF 11/20/2008] DEV: All e-mails are sent to developer during test mode
		If CONST_SEND_MAIL_TO_DEV Then
    		locstrTo = CONST_DEVELOPER_EMAIL ' Set in appglobal.asp
    	End If
       
        Set objCDOMail = CreateObject("CDO.Message")
        objCDOMail.Configuration = cdoConfig 'added by [MFILLAST 08-2006]
        objCDOMail.From = locstrFrom
        objCDOMail.To = locstrTo
        objCDOMail.Subject = locstrSubject
        objCDOMail.HTMLBody = trim(locstrBody)
        objCDOMail.AddAttachment server.mappath(CONST_APPLICATION_PATH & "/common/images/email.png"), "email.png"
        
        objCDOMail.Send
         
        If Err.number <> 0 Then
            locReturnValue = False
        End If
        
        set objCDOMail = nothing
        Set cdoConfig = Nothing 'added by [MFILLAST 08-2006]
        
        mSendEmail = locReturnValue
    end function
    
    function updateEmployees(ByRef locobjEEtoView)
		'local connections
            Dim locCmd
            Dim locCmdRevoke
            Dim locRS
            Dim locCmd2
			Dim locRS2
			Dim locCmd3
			Dim locRS3
			Dim locCmd4
			Dim locRS4
			Dim locCmd5
			
            'wds connection
            Dim wdsRS
			Dim wdsConnection
			Dim locCmdwds
			Dim wdsRS2
			Dim wdsConnection2
			Dim locCmdwds2
			
			'vars
			Dim insertRequired 
			Dim updateRequired 
			Dim condition
			Dim condition2
            Dim RegionCode
            Dim EmpTypeCode
					
            '*** If we have no WWID set up, and no IDSID set up we can't find our ee details - so exit the routine.
            If locObjEEtoView.WWID = "" then
                exit function
            End if
            
            'condition = "NextLevelWWID = '10705764';" 
            condition = " ""NextLevelWwid"" = '" & locobjEEtoView.WWID & "' "  
            
            'connection to WDS 
			Set wdsConnection = Server.CreateObject("ADODB.Connection")
			wdsConnection.ConnectionString = CONST_WDS_CONNECTION_STRING
			wdsConnection.Open
			Set locCmdwds = Server.CreateObject("ADODB.Command")
			Set locCmdwds.ActiveConnection = wdsConnection
			locCmdwds.CommandText = " SELECT * FROM ""WorkerSnapshotConfidentialV2"" where " & condition

			set wdsRS = locCmdwds.Execute
			'TODO check result
			
			
			do while not wdsRS.eof
			
				'create the connection
				Set locCmd2 = Server.CreateObject("ADODB.Command")
				Set locCmd2.ActiveConnection = glbConnection
				locCmd2.CommandType = adCmdStoredProc
				locCmd2.CommandText = "usp_getLastUpdateWdsCopy_from_wwid"
                locCmd2.Parameters.Append locCmd2.CreateParameter("strWWID", adWChar, adParamInput, 8, wdsRS("WWID"))
                
                Set locRS2 = locCmd2.Execute
                
                updateRequired = false
				insertRequired = false
				
				if isnull(locRS2) or locRS2.eof then ' no entry found
					insertRequired = true
				else
					updateRequired = true
				end if
				
				 locRS2.Close
				Set locRS2 = nothing
				Set locCmd2 = nothing
					
					 if insertRequired or updateRequired then
					 
						'create mssql connection
						Set locCmd = Server.CreateObject("ADODB.Command")
						Set locCmd.ActiveConnection = glbConnection
						locCmd.CommandType = adCmdStoredProc
				
						if insertRequired then
							locCmd.CommandText = "usp_insert_wds_copy"
						Else 'update
							locCmd.CommandText = "usp_update_wds_copy"      
						end If
						
						locCmd.Parameters.Append locCmd.CreateParameter("WWID", adWChar, adParamInput, 8 , wdsRS("WWID"))
                        mDebugPrint "wds(WWID) = '" & wdsRS("WWID") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("FirstNm", adWChar, adParamInput, 30 , Left(wdsRS("FirstNm"),30))
                        mDebugPrint "wds(FirstNm) = '" & wdsRS("FirstNm") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("LastNm", adWChar, adParamInput, 30 , Left(wdsRS("LastNm"),30))
                        mDebugPrint "wds(LastNm) = '" & wdsRS("LastNm") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("Idsid", adWChar, adParamInput, 8 , wdsRS("Idsid"))
                        mDebugPrint "wds(Idsid) = '" & wdsRS("idsid") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("EmployeeStatusCd", adWChar, adParamInput, 1 , wdsRS("EmployeeStatusCd"))
                        mDebugPrint "wds(EmployeeStatusCd) = '" & wdsRS("EmployeeStatusCd") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("CorporateEmailTxt", adWChar, adParamInput, 80 , Left(wdsRS("CorporateEmailTxt"),80))
                        mDebugPrint "wds(CorporateEmailTxt) = '" & wdsRS("CorporateEmailTxt") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("CompanyCd", adWChar, adParamInput, 3 , wdsRS("CompanyCd"))
                        mDebugPrint "wds(CompanyCd) = '" & wdsRS("CompanyCd") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("MailStopTxt", adWChar, adParamInput, 50 , Left(wdsRS("MailStopTxt"),50))
                        mDebugPrint "wds(MailStopTxt) = '" & wdsRS("MailStopTxt") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("NextLevelNm", adWChar, adParamInput, 50 , Left(wdsRS("NextLevelNm"),50))
                        mDebugPrint "wds(NextLevelNm) = '" & wdsRS("NextLevelNm") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("NextLevelWWID", adWChar, adParamInput, 8 , wdsRS("NextLevelWwid"))
                        mDebugPrint "wds(NextLevelWWID) = '" & wdsRS("NextLevelWwid") & "'<br>"

                             'deprecated locCmd.Parameters.Append locCmd.CreateParameter("RegionCode", adWChar, adParamInput, 4 , wdsRS("RegionCode"))
                
                        mDebugPrint "wds(GeographyCd) = '" & wdsRS("GeographyCd") & "'<br>"
                        mDebugPrint "wds(CompanyCountryCd) = '" & wdsRS("CompanyCountryCd") & "'<br>"
                            if wdsRS("GeographyCd")  = "AMER" AND ( not wdsRS("CompanyCountryCd") = "USA" AND not wdsRS("CompanyCountryCd") = "CAN") THEN 
                                RegionCode = "LAT"
                            elseif wdsRS("GeographyCd")  = "AMER" AND (wdsRS("CompanyCountryCd") = "USA" OR wdsRS("CompanyCountryCd") = "US" ) THEN 
                                RegionCode = "LAT"
                            elseif wdsRS("GeographyCd")  = "AMER" AND wdsRS("CompanyCountryCd") = "USA" THEN
                                 RegionCode = "US"
                            elseif wdsRS("GeographyCd")  = "AMER" AND wdsRS("CompanyCountryCd") = "CAN" THEN
                                 RegionCode = "CNDA"
                            elseif wdsRS("GeographyCd")  = "APAC" THEN
                                 RegionCode = "APAC"
                            elseif wdsRS("GeographyCd")  = "EMEA" AND (not wdsRS("CompanyCountryCd") = "ISR" AND not wdsRS("CompanyCountryCd") = "IRL")  THEN
                                 RegionCode = "EMEA"
                            elseif wdsRS("CompanyCountryCd") = "ISR" THEN
                                 RegionCode = "ISR"  
                            elseif wdsRS("CompanyCountryCd") = "IRL" THEN
                                 RegionCode = "IRE"
                            else
                                RegionCode = NULL
                            end if
                        locCmd.Parameters.Append locCmd.CreateParameter("RegionCode", adWChar, adParamInput, 4 , RegionCode)
                        mDebugPrint "wds(RegionCode) = '" & RegionCode & "'<br>"

                        locCmd.Parameters.Append locCmd.CreateParameter("WorkLocationSiteCd", adWChar, adParamInput, 3 , wdsRS("WorkLocationSiteCd"))
                        mDebugPrint "wds(WorkLocationSiteCd) = '" & wdsRS("WorkLocationSiteCd") & "'<br>"

                        locCmd.Parameters.Append locCmd.CreateParameter("EndDt", adDBTimeStamp, adParamInput, , wdsRS("EndDt"))
                        mDebugPrint "wds(EndDt) = '" & wdsRS("EndDt") & "'<br>"

                        locCmd.Parameters.Append locCmd.CreateParameter("StartDt", adDBTimeStamp, adParamInput,, wdsRS("StartDt")) 
                        mDebugPrint "wds(StartDt) = '" & wdsRS("StartDt") & "'<br>"

                             'deprecated locCmd.Parameters.Append locCmd.CreateParameter("EmpTypeCode", adWChar, adParamInput, 3, wdsRS("EmpTypeCode")) 
                
                        mDebugPrint "wds(EmployeeClassCd) = '" & wdsRS("EmployeeClassCd") & "'<br>"
                        mDebugPrint "wds(PersonRoleTypeCd) = '" & wdsRS("PersonRoleTypeCd") & "'<br>"
                        if wdsRS("EmployeeClassCd")  = "E" OR wdsRS("EmployeeClassCd")  = "C" OR wdsRS("EmployeeClassCd")  = "T" OR wdsRS("EmployeeClassCd")  = "R" THEN 
                            EmpTypeCode = "REG"
                        elseif wdsRS("EmployeeClassCd")  = "S" THEN
                            EmpTypeCode = "SH"
                        elseif wdsRS("EmployeeClassCd")  = "I" THEN
                            EmpTypeCode = "ICE"
                        elseif wdsRS("EmployeeClassCd")  = "" OR wdsRS("EmployeeClassCd")  = NULL  THEN
                            if wdsRS("PersonRoleTypeCd")  = "IDP" THEN
                                EmpTypeCode = "IC"
                            elseif wdsRS("PersonRoleTypeCd")  = "OSR" THEN
                                EmpTypeCode = "OS"
                            elseif wdsRS("PersonRoleTypeCd")  = "SAG" THEN
                                EmpTypeCode = "SA"
                            elseif wdsRS("PersonRoleTypeCd")  = "POI" THEN
                                EmpTypeCode = "PV"
                            else
                                EmpTypeCode = "TMP"
                            end if   
                        else
                            EmpTypeCode = ""
                        end if

                        locCmd.Parameters.Append locCmd.CreateParameter("EmpTypeCode", adWChar, adParamInput, 4 , EmpTypeCode)
                        mDebugPrint "wds(EmpTypeCode) = '" & EmpTypeCode & "'<br>"

                        locCmd.Parameters.Append locCmd.CreateParameter("OriginalStartDt", adDBTimeStamp, adParamInput,, wdsRS("OriginalStartDt")) 
                        mDebugPrint "wds(OriginalStartDt) = '" & wdsRS("OriginalStartDt") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("FullTmPartTmCd", adWChar, adParamInput, 2 , wdsRS("FullTmPartTmCd"))
                        mDebugPrint "wds(FullTmPartTmCd) = '" & wdsRS("FullTmPartTmCd") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("FLSACd", adWChar, adParamInput, 1, wdsRS("FLSACd")) 
                        mDebugPrint "wds(FLSACd) = '" & wdsRS("FLSACd") & "'<br>"
                        locCmd.Parameters.Append locCmd.CreateParameter("DepartmentNm", adWChar, adParamInput, 50, Left(wdsRS("DepartmentNm"),50))
                        mDebugPrint "wds(DepartmentNm) = '" & wdsRS("DepartmentNm") & "'<br>"
						
						locCmd.Execute
                
						Set locCmd = nothing
					 
					 
					 end if
					 
			
			wdsRS.movenext
			loop
			
			wdsRS.close
            Set wdsRS = nothing
            Set locCmdwds = nothing
            
            
            'create the connection
			Set locCmd3 = Server.CreateObject("ADODB.Command")
			Set locCmd3.ActiveConnection = glbConnection
			locCmd3.CommandType = adCmdStoredProc
			locCmd3.CommandText = "dbo.get_emps" 	' get all employees of a manager using the manager's WWID
            locCmd3.Parameters.Append locCmd3.CreateParameter("@vWWID", adWChar, adParamInput, 8, locobjEEtoView.WWID)
            
            Set locRS3 = locCmd3.Execute
                
            do while not locRS3.eof
        
				
				updateRequired = true ' MOF: EVERYONE IS UPDATED..... CHECKING THE LAST UPDATED TIME is pointless
				
					if updateRequired then
					    
                        condition2 = " ""WWID"" = '" & locRS3("WWID") & "' "  
						updateOrInsert updateRequired,false,condition2
					
					end if
            
            locRS3.movenext
            loop
            
            response.Redirect "teamleavesummary.asp"
    end function


        function updateOrInsert(ByVal updateRequired, ByVal insertRequired, ByVal condition)
                'local connections
                Dim locCmd2
            
                'wds connection
                Dim wdsRS
                Dim wdsConnection
                Dim locCmdwds

                'vars
                Dim RegionCode
                Dim EmpTypeCode
                
                'connection to WDS 
                Set wdsConnection = Server.CreateObject("ADODB.Connection")
                wdsConnection.ConnectionString = CONST_WDS_CONNECTION_STRING
                wdsConnection.Open
                mDebugPrint wdsConnection & "<br>"
                Set locCmdwds = Server.CreateObject("ADODB.Command")
                Set locCmdwds.ActiveConnection = wdsConnection

                locCmdwds.CommandText = " SELECT * FROM ""WorkerSnapshotConfidentialV2"" where " & condition & " order by ""EndDt"" desc"
                mDebugPrint locCmdwds.CommandText & "<br>"
                mDebugPrint locCmdwds.State & "<br>"
                set wdsRS = locCmdwds.Execute

                'TODO check result
                mDebugPrint "wdsRS.EOF = " & wdsRS.EOF & "<br>"
                mDebugPrint "wds Command Text = '" & locCmdwds.CommandText & "'<br>"
										
                'create mssql connection
                Set locCmd2 = Server.CreateObject("ADODB.Command")
                Set locCmd2.ActiveConnection = glbConnection
                locCmd2.CommandType = adCmdStoredProc

                

                locCmd2.Parameters.Append locCmd2.CreateParameter("WWID", adWChar, adParamInput, 8 , wdsRS("WWID"))
                mDebugPrint "wds(WWID) = '" & wdsRS("WWID") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("FirstNm", adWChar, adParamInput, 30 , Left(wdsRS("FirstNm"),30))
                mDebugPrint "wds(FirstNm) = '" & wdsRS("FirstNm") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("LastNm", adWChar, adParamInput, 30 , Left(wdsRS("LastNm"),30))
                mDebugPrint "wds(LastNm) = '" & wdsRS("LastNm") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("Idsid", adWChar, adParamInput, 8 , wdsRS("Idsid"))
                mDebugPrint "wds(Idsid) = '" & wdsRS("idsid") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("EmployeeStatusCd", adWChar, adParamInput, 1 , wdsRS("EmployeeStatusCd"))
                mDebugPrint "wds(EmployeeStatusCd) = '" & wdsRS("EmployeeStatusCd") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("CorporateEmailTxt", adWChar, adParamInput, 80 , Left(wdsRS("CorporateEmailTxt"),80))
                mDebugPrint "wds(CorporateEmailTxt) = '" & wdsRS("CorporateEmailTxt") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("CompanyCd", adWChar, adParamInput, 3 , wdsRS("CompanyCd"))
                mDebugPrint "wds(CompanyCd) = '" & wdsRS("CompanyCd") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("MailStopTxt", adWChar, adParamInput, 50 , Left(wdsRS("MailStopTxt"),50))
                mDebugPrint "wds(MailStopTxt) = '" & wdsRS("MailStopTxt") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("NextLevelNm", adWChar, adParamInput, 50 , Left(wdsRS("NextLevelNm"),50))
                mDebugPrint "wds(NextLevelNm) = '" & wdsRS("NextLevelNm") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("NextLevelWWID", adWChar, adParamInput, 8 , wdsRS("NextLevelWwid"))
                mDebugPrint "wds(NextLevelWWID) = '" & wdsRS("NextLevelWwid") & "'<br>"

                     'deprecated locCmd2.Parameters.Append locCmd2.CreateParameter("RegionCode", adWChar, adParamInput, 4 , wdsRS("RegionCode"))
                
                mDebugPrint "wds(GeographyCd) = '" & wdsRS("GeographyCd") & "'<br>"
                mDebugPrint "wds(CompanyCountryCd) = '" & wdsRS("CompanyCountryCd") & "'<br>"
                if wdsRS("GeographyCd")  = "AMER" AND ( not wdsRS("CompanyCountryCd") = "USA" AND not wdsRS("CompanyCountryCd") = "CAN") THEN 
                    RegionCode = "LAT"
                elseif wdsRS("GeographyCd")  = "AMER" AND (wdsRS("CompanyCountryCd") = "USA" OR wdsRS("CompanyCountryCd") = "US" ) THEN 
                    RegionCode = "LAT"
                elseif wdsRS("GeographyCd")  = "AMER" AND wdsRS("CompanyCountryCd") = "USA" THEN
                     RegionCode = "US"
                elseif wdsRS("GeographyCd")  = "AMER" AND wdsRS("CompanyCountryCd") = "CAN" THEN
                     RegionCode = "CNDA"
                elseif wdsRS("GeographyCd")  = "APAC" THEN
                     RegionCode = "APAC"
                elseif wdsRS("GeographyCd")  = "EMEA" AND (not wdsRS("CompanyCountryCd") = "ISR" AND not wdsRS("CompanyCountryCd") = "IRL")  THEN
                     RegionCode = "EMEA"
                elseif wdsRS("CompanyCountryCd") = "ISR" THEN
                     RegionCode = "ISR"  
                elseif wdsRS("CompanyCountryCd") = "IRL" THEN
                     RegionCode = "IRE"
                else
                    RegionCode = NULL
                end if
                locCmd2.Parameters.Append locCmd2.CreateParameter("RegionCode", adWChar, adParamInput, 4 , RegionCode)
                mDebugPrint "wds(RegionCode) = '" & RegionCode & "'<br>"

                locCmd2.Parameters.Append locCmd2.CreateParameter("WorkLocationSiteCd", adWChar, adParamInput, 3 , wdsRS("WorkLocationSiteCd"))
                mDebugPrint "wds(WorkLocationSiteCd) = '" & wdsRS("WorkLocationSiteCd") & "'<br>"

                locCmd2.Parameters.Append locCmd2.CreateParameter("EndDt", adDBTimeStamp, adParamInput, , wdsRS("EndDt"))
                mDebugPrint "wds(EndDt) = '" & wdsRS("EndDt") & "'<br>"

                locCmd2.Parameters.Append locCmd2.CreateParameter("StartDt", adDBTimeStamp, adParamInput,, wdsRS("StartDt")) 
                mDebugPrint "wds(StartDt) = '" & wdsRS("StartDt") & "'<br>"

                     'deprecated locCmd2.Parameters.Append locCmd2.CreateParameter("EmpTypeCode", adWChar, adParamInput, 3, wdsRS("EmpTypeCode")) 
                
                mDebugPrint "wds(EmployeeClassCd) = '" & wdsRS("EmployeeClassCd") & "'<br>"
                mDebugPrint "wds(PersonRoleTypeCd) = '" & wdsRS("PersonRoleTypeCd") & "'<br>"
                if wdsRS("EmployeeClassCd")  = "E" OR wdsRS("EmployeeClassCd")  = "C" OR wdsRS("EmployeeClassCd")  = "T" OR wdsRS("EmployeeClassCd")  = "R" THEN 
                    EmpTypeCode = "REG"
                elseif wdsRS("EmployeeClassCd")  = "S" THEN
                    EmpTypeCode = "SH"
                elseif wdsRS("EmployeeClassCd")  = "I" THEN
                    EmpTypeCode = "ICE"
                elseif wdsRS("EmployeeClassCd")  = "" OR wdsRS("EmployeeClassCd")  = NULL  THEN
                    if wdsRS("PersonRoleTypeCd")  = "IDP" THEN
                        EmpTypeCode = "IC"
                    elseif wdsRS("PersonRoleTypeCd")  = "OSR" THEN
                        EmpTypeCode = "OS"
                    elseif wdsRS("PersonRoleTypeCd")  = "SAG" THEN
                        EmpTypeCode = "SA"
                    elseif wdsRS("PersonRoleTypeCd")  = "POI" THEN
                        EmpTypeCode = "PV"
                    else
                        EmpTypeCode = "TMP"
                    end if   
                else
                    EmpTypeCode = ""
                end if

                locCmd2.Parameters.Append locCmd2.CreateParameter("EmpTypeCode", adWChar, adParamInput, 4 , EmpTypeCode)
                mDebugPrint "wds(EmpTypeCode) = '" & EmpTypeCode & "'<br>"

                locCmd2.Parameters.Append locCmd2.CreateParameter("OriginalStartDt", adDBTimeStamp, adParamInput,, wdsRS("OriginalStartDt")) 
                mDebugPrint "wds(OriginalStartDt) = '" & wdsRS("OriginalStartDt") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("FullTmPartTmCd", adWChar, adParamInput, 2 , wdsRS("FullTmPartTmCd"))
                mDebugPrint "wds(FullTmPartTmCd) = '" & wdsRS("FullTmPartTmCd") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("FLSACd", adWChar, adParamInput, 1, wdsRS("FLSACd")) 
                mDebugPrint "wds(FLSACd) = '" & wdsRS("FLSACd") & "'<br>"
                locCmd2.Parameters.Append locCmd2.CreateParameter("DepartmentNm", adWChar, adParamInput, 50, Left(wdsRS("DepartmentNm"),50))
                mDebugPrint "wds(DepartmentNm) = '" & wdsRS("DepartmentNm") & "'<br>"
	            
                if wdsRS("StartDt") = null or wdsRS("OriginalStartDt") = null then
                    insertRequired = False
                    updateRequired = False
                end if

                if insertRequired then
                    locCmd2.CommandText = "usp_insert_wds_copy"						
                    ' Create a Local Record of the Employee
                    Dim user
                    Set user = new cObjUser
                    user.CreateLocalUserRecord 
                    locCmd2.Execute
                Else 'update
                    locCmd2.CommandText = "usp_update_wds_copy"
                    locCmd2.Execute      
                end If

                '   end if ' [MOF] End check on Record Set

                wdsRS.close
                Set wdsRS = nothing
                Set locCmd2 = nothing
                wdsConnection.Close
                Set wdsConnection = nothing

    end function

            function RenewDbCleanDb()
                Dim locCmd3
                Dim locRS3
                Dim locCmd4
                Dim locRS4
                Dim condition
                Dim fld

                Set locCmd4 = Server.CreateObject("ADODB.Command")
                Set locCmd4.ActiveConnection = glbConnection
                locCmd4.CommandType = adCmdStoredProc
                
                mDebugPrint "Cleaning up"
                locCmd4.CommandText = "clean_up_db"
                Set locRS4 = locCmd4.Execute

                Set locCmd3 = Server.CreateObject("ADODB.Command")
                Set locCmd3.ActiveConnection = glbConnection
                locCmd3.CommandType = adCmdStoredProc
                
                mDebugPrint "Updating all records"
                locCmd3.CommandText = "getWWID"
                Set locRS3 = locCmd3.Execute
                
                While not locRS3.eof
                    for each fld in locRS3.Fields
                        mDebugPrint fld & "<br>"
                        condition = """WWID"" = '" & fld & "'"
                        mDebugPrint condition & "<br>"
                        updateOrInsert true,false,condition
                    Next
			            
			        locRS3.movenext
		        Wend
                
                locRS3.Close
                Set locRS3 = nothing           
                Set locCmd3 = nothing

                Set locRS4 = locCmd4.Execute

                locRS4.Close
                Set locRS4 = nothing           
                Set locCmd4 = nothing
                
    end function
    
        %>