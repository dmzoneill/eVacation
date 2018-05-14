<%
	'*************************************************
	'PAGE STRUCTURE SUB ROUTINES / FUNCTIONS :
	'=================================================
	'mWriteAdminCarryOver		(ByRef locobjUser)
	'mWriteAdminELP				(ByRef locobjUser, ByRef locobjActiveELP, ByRef locobjELPRelief)
	'mWriteAdminLeavePeriod		(ByRef locobjUser, ByRef locobjLeaveRequest, ByVal locblnValidateRequest)
	'mWriteAdminMenu			(ByRef locobjSelectedUser, ByVal locstrCurrentSelection)
	'mWriteAdminSelectUserForm	(ByRef locobjUser, ByVal locblnValidateUser)
	'mWriteAdminUserProfileForm	(ByRef locobjUser, ByVal locblnValidateProfile)
	'mWriteApproveLeaveRequest	(ByRef locobjLeaveRequest, ByVal locblnValidateApproval)
	'mWriteApproveRequests		(ByRef locobjUser)
	'mWriteCancelLeavePeriodRequestForm		(ByRef locobjUser, ByRef locobjLeavePeriod, ByVal locblnValidateRequest)
	'mWriteDelegateForm			(ByRef locobjUser, ByVal locblnValidate)
	'mWriteELPDetails			(ByRef locobjELP, ByVal locblnAllowEdit)
	'mWriteELPEditForm			(ByVal locobjELPInstance)
	'mWriteFormError			(ByVal loclngErrorType, ByVal locstrMessage)
	'mWriteGeneralError			(ByVal locstrMessage, ByVal locblnBackButtonMessage)
	'mWriteHomePage				()
	'mWriteHMTLTop				(ByVal locstrTitle)
	'	openPopWin				CS JavaScript (winURL, winWidth, winHeight, winFeatures, winLeft, winTop)
	'	closePopWin				CS JavaScript ()
	'	getLocation				CS JavaScript (winWidth, winHeight, winLeft, winTop)
	'mWriteLeaveRequestForm		(ByRef locobjUser, ByRef locobjLeaveRequest, ByVal locblnValidateRequest)
	'mWriteCompTimeInstanceDetails      (ByVal loclngCompTimeID)
	'mEraseCompTimeInstanceDetails      (ByVal loclngCompTimeID)

	'mWriteLeaveRequests		(ByRef locobjLeaveRequests, ByVal locblnAdminView)
	'mWriteCompTimeDetails
	'mWriteLeaveSummary			(ByRef locobjUser)
	'mWriteNavBar				(ByVal locstrTabName)
	'mWritePageFooter			()
	'mWritePublicHolidays		()
	'mWriteReportForm			()
	'mWriteSelectYearOptions	(ByVal loclngYear)
	'mWriteTeamLeaveSummary		(ByRef locobjUser)
	'mWriteTimeOption			(ByVal locstrFieldName, ByVal locstrTime, ByVal locblnDisabled)
	'mWriteUserError			(ByVal loclngErrorCode)
	'mWriteViewingEmployee		(ByRef locobjUser)
	'mWriteELPVacationDetails	(ByRef locobjLeavePeriod, ByRef locobjELPInstance)
	'mWriteMaturedELPBankedDays (ByRef locobjELP)
	'mWriteMaturedELPDetails	(ByRef locobjELPInstance, ByVal Admin View)
	'mWriteELPLeaveRequestForm	()
	'mWriteELPAdminLeaveRequestForm ()
	'mWriteCalendarHolidayDisplay(ByRef objEEtoView)
	'mWriteAddHolidayForm		()
	'mWriteRemoveHolidayForm	()
	'mWriteAddLeaveType			()
	'mWriteRemoveLeaveType		()
	'mWriteSqlLinks				()
	'mWriteHomePage				()
	'
	'***************************************************************


	'*** WRITE ADMIN CARRY OVER ***
	function mWriteAdminCarryOver(ByRef locobjUser, ByRef locobjCarryOver, ByVal locblnValidateCarryOver)
		Dim loclngCounter
		Dim loclngCount
		Dim loclngYear
		
		response.write "<table class='pageContentHeader'><tr><td><div class='th'>Carry over / Exemptions</div></td></tr></table>"
		
		With locobjUser
			response.write "<br><table class='pageContentTable' style='width:600px'>"
				response.write "<tr>"
					'*** CARRY OVER EOY ***
					response.write "<td valign=top>"
						response.write "<span class='th'>"
						response.write "Carry Over EOY"
						response.write "</span><table>"
							response.write "<tr>"
								response.write "<th style='width:50px'>"
								response.write "</th>"
								response.write "<th>"
									response.write "Year"
								response.write "</th>"
								response.write "<th>"
									response.write "Days"
								response.write "</th>"
							response.write "</tr>"
							With .CarryOversEOY
								loclngCounter = 0
								loclngCount = .Count
								loclngYear = DatePart("yyyy",date())
								While loclngCounter < loclngCount
									loclngCounter = loclngCounter +1
									if .Item(loclngCounter).Year > loclngYear then
										loclngYear = .Item(loclngCounter).Year
									end if
								Wend
							End With
							For loclngCounter = loclngYear to CONST_FIRST_YEAR_SYSTEM_ACTIVE Step - 1
								With .CarryOverEOYForYear(loclngCounter)
									response.write "<tr>"
										response.write "<td>"
											response.write "<a href='#' onclick=""javascript:alert(this.title);"" title=""Entered By: "
												response.write mHTMLEncode(.EnteredBy.FullName)
												response.write chr(13) & chr(13) & "Entered On: "
												response.write mFormatDate(.DateEntered,"medium with day")
												response.write chr(13) & chr(13) & "Comments: "
												response.write mHTMLEncode(.Comments)
												response.write chr(13)
											response.write """><img src=""" & CONST_APPLICATION_PATH & "/common/images/info.png"" width=16 height=16 style='border:0px' alt=''></a>"
										response.write "</td>"
										response.write "<td>"
											response.write loclngCounter
										response.write "</td>"	
										response.write "<td>"
											response.write .Days
										response.write "</td>"
									response.write "</tr>"
								End With
							next
						response.write "</table>"
					response.write "</td><td style='width:100px'><br></td>"
					
					'*** CARRY OVER PRE-ARRANGED ***
					response.write "<td valign=top>"
						response.write "<span class='th'>Exceptions</span>"
						response.write "<table>"
							response.write "<tr>"								
								response.write "<th style='width:50px'>"
								response.write "</th>"
								response.write "<th>"
									response.write "Year"
								response.write "</th>"
								response.write "<th>"
									response.write "Days"
								response.write "</th>"
							response.write "</tr>"
							With .CarryOversPreArranged
								loclngCounter = 0
								loclngCount = .Count
								While loclngCounter < loclngCount
									loclngCounter = loclngCounter +1
									if .Item(loclngCounter).Year > loclngYear then
										loclngYear = .Item(loclngCounter).Year
									end if
								Wend
							End With
							For loclngCounter = loclngYear to CONST_FIRST_YEAR_SYSTEM_ACTIVE Step - 1
								With .CarryOverPreArrangedForYear(loclngCounter)
									response.write "<tr>"
										response.write "<td align=left>"
											response.write "<a href='#' onclick=""javascript:alert(this.title);"" title=""Entered By: "
												response.write mHTMLEncode(.EnteredBy.FullName)
												response.write chr(13) & chr(13) & "Entered On: "
												response.write mFormatDate(.DateEntered,"medium with day")
												response.write chr(13) & chr(13) & "Comments: "
												response.write mHTMLEncode(.Comments)
												response.write chr(13)
											response.write """><img src=""" & CONST_APPLICATION_PATH & "/common/images/info.png"" width=16 height=16 style='border:0px' alt=''></a>"
										response.write "</td>"
										response.write "<td>"
											response.write loclngCounter
										response.write "</td>"	
										response.write "<td>"
											response.write .Days
										response.write "</td>"
									response.write "</tr>"
								End With
							next
						response.write "</table>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"
				
			response.write "<table class='pageContentHeader'><tr><td><div class='th'>Update Carry Over/Exceptions</div></td></tr></table>"
			
				If locblnValidateCarryOver then
					response.write "<table class='pageContentTable'>"
					response.write "<tr>"
						response.write "<td style='text-align:center' colspan=2>"
							If not locobjCarryOver.AdminCarryOverFormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjCarryOver.AdminCarryOverFormErrorMessage
							End If
						response.write "</td>"
					response.write "</tr>"
					response.write "</table>"
				End If

			response.write "<table class='pageContentTable'>"
				response.write "<form name=frmCarryOver action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"
					response.write "<tr>"
		
						'**** START HIDDEN FIELD VALUES ****
						response.write "<input type=hidden name=m value=co>"
						response.write "<input type=hidden name=formname value=frmCarryOver>"
						response.write "<input type=hidden name=ee value="""
							response.write .WWID
						response.write """>"
						response.write "<input type=hidden name=itemid value="""
							response.write locobjCarryOver.ID
						response.write """>"
						'**** END HIDDEN FIELD VALUES ****
						
						response.write "<td colspan=2>"
							response.write "<div class='th' style='margin-bottom:10px'>Type </div>"
							response.write " <select class='basicselect' name=fldcbotype>"
								response.write "<option value=""EOY"" "
									If not locobjCarryOver.IsPreArranged then
										response.write "selected"
									End If
								response.write ">EOY"
								response.write "<option value=""Pre-Arranged"""
									If locobjCarryOver.IsPreArranged then
										response.write "selected"
									End If
								response.write ">Exception"
							response.write "</select>"
							response.write "<div class='th' style='margin-bottom:10px'>Year</div> "
							response.write " <select class='basicselect' name='fldlngYear'><option value=''>Select year</option>"
							
							Dim i
							
							For i = 2000 to 2030             
								response.write "<option value='" & i & "'"
								if Cint( mHTMLEncode(locobjCarryOver.Year)) = i then
									response.write " selected='selected'"
								End If
								response.write ">" & i & "</option>"
							Next
							
							response.write "</select>"
							response.write "<div class='th' style='margin-bottom:10px'>Days</div> "
							response.write "<input name=flddbldays type=text maxlength=5 size=5 value="""
								response.write mHTMLEncode(locobjCarryOver.Days)
							response.write """>"
						response.write "</td>"
					response.write "</tr>"
					response.write "<tr>"
						response.write "<td colspan=2>"
							response.write "<div class='th' style='margin-bottom:10px'>Comments</div>"
							response.write "<textarea name=fldstrcomments rows=3 cols=40>"
								response.write mHTMLEncode(locobjCarryOver.Comments)
							response.write "</textarea><br>"
						response.write "</td>"
					response.write "</tr>"
					response.write "<tr>"
						response.write "<td colspan=2>"
							response.write "<input type=submit value=""Update Carry Over/Exceptions""><br>"
							response.write "<br>"
						response.write "</td>"
					response.write "</tr>"
				response.write "</form>"
			end with					
		response.write "</table>"
		response.write ""
	end function

	'*** WRITE ADMIN ELP *** @@ELPHERE
	function mWriteAdminELP(ByRef locobjUser, ByRef locobjActiveELP, ByRef locobjELPRelief, ByVal locblnValidateActivateELP)
		
			    				
		Dim loclngYear
		Dim loclngEndYear
		Dim locarrAddReliefYears()
		Dim locarrRemoveReliefYears()
		Dim loclngCounter
		Dim locobjELPLeaveRequest
		Dim locblnValidateRequest
		Dim loclngSaveResult
		Dim locobjELPLeavePeriod
		Dim loclngResult
		Dim locblnDisplaySummary
		Dim locblnDisplayCancelRequestForm
		
		
		Set locobjELPLeaveRequest = new cObjLeavePeriod
	
		locblnDisplaySummary = True
		
		If request.form("formname") = "frmRequestELPLeave" then
			locblnValidateRequest = True
			locobjELPLeaveRequest.LoadNewRequestFromForm
		Else
			locblnValidateRequest = False
		End If
		
		if locblnValidateRequest then
			if locobjELPLeaveRequest.FormIsValid then
				if locobjELPLeaveRequest.NewRequestIsValid then
					locobjELPLeaveRequest.IsAdminUpdating = true
					loclngSaveResult = locobjELPLeaveRequest.Save
				
			'	response.write "Test Str304 " & locobjUser &  " - " & locobjActiveELP	
				
					If loclngSaveResult = 0 then
						mWriteGeneralError "Sorry - an error has occurred while attempting to save your leave request.<br>" & _
							"There may be a problem with the network, or the database server may be experiencing difficulties.", False
					else
						response.write "<center>"
							response.write "<br>"
							response.write "<b>Your ELP leave request has been raised successfully.</b><br>"
							response.write "<br>"
						response.write "</center>"
						Set locobjELPLeaveRequest = nothing			
						Set locobjELPLeaveRequest = new cObjLeavePeriod
						
						locblnValidateRequest = False
					
					end if
				end if
			end if
		end if
	
	
		'*** For Cancelling a leave period (ELP Leave period in this case!)
		If request.querystring("strMode") = "cl" then
			set locobjELPLeavePeriod = new cObjLeavePeriod
			locobjELPLeavePeriod.ID = lngItemID
			If not locobjELPLeavePeriod.IsActive then
				mWriteGeneralError "Sorry - the leave period you are attempting to cancel is not active.", False
			Else
						
					loclngResult = locobjELPLeavePeriod.Delete
					if loclngResult = 1 then
						mWriteGeneralError "Sorry - an error has occurred while attempting to cancel the leave period.<br>" & _
									"There may be a problem with the network, or the database server may be experiencing difficulties.", False
					Else
						response.write "<center>"
							response.write "<br>"
							response.write "<b>The leave period has been cancelled successfully.</b><br>"
							response.write "<br>"
						response.write "</center>"
					End If
				'End If
			End If
		End If

		If locblnDisplaySummary then
			
			'***Mature ELP
			
			If objEEtoView.AnnualVacation.HasMaturedELP then
				mWriteMaturedELPDetails objEEtoView.AnnualVacation.ELPMatured, true
				If not objEEtoView.AnnualVacation.ELPMatured.IsExpired then
					mWriteELPAdminLeaveRequestForm objEEtoView, locobjELPLeaveRequest, locblnValidateRequest
				End If
			End If
			
			'**ELP Vacation Details
			If objEEtoView.AnnualVacation.ELPUsed.IsUsed then
				mWriteELPVacationDetails objEEtoView, objEEtoView.AnnualVacation.ELPUsed, true
			End If
					
			If locobjUser.AnnualVacation.HasActiveELP then
				
				mWriteELPDetails locobjActiveELP, True
				
				With locobjActiveELP
	
					loclngYear = DatePart("yyyy",.DateActivated) + 1
					loclngEndYear = DatePart("yyyy",.MaturityDate)
					Redim locarrRemoveReliefYears(0)
					Redim locarrAddReliefYears(0)
	
					While loclngYear <= loclngEndYear
						If .ReliefInYear(loclngYear) then
							Redim preserve locarrRemoveReliefYears(ubound(locarrRemoveReliefYears)+1)
							locarrRemoveReliefYears(ubound(locarrRemoveReliefYears)) = loclngYear
						Else
							Redim preserve locarrAddReliefYears(ubound(locarrAddReliefYears)+1)
							locarrAddReliefYears(ubound(locarrAddReliefYears)) = loclngYear
						End If
						loclngYear = loclngYear + 1
					Wend
					
					response.write "<table class='pageContentTable'>"
		
						response.write "<form name=frmEditELPRelief action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"
							'**** START HIDDEN FIELD VALUES ****
							response.write "<input type=hidden name=formname value=frmEditELPRelief>"
							response.write "<input type=hidden name=ee value="""
								response.write locobjUser.WWID
							response.write """>"
							response.write "<input type=hidden name=m value="""
								response.write strMode
							response.write """>"
							response.write "<input type=hidden name=elpid value="""
								response.write .ELPID
							response.write """>"
							'**** END HIDDEN FIELD VALUES ****
		
							response.write "<tr>"
								response.write "<td style='width:50px'>"
									response.write ""
								response.write "</td>"
								response.write "<td colspan=4 class=txttitle width=150 >"
									response.write "Active ELP Reliefs:"
								response.write "</td>"
								response.write "<td>"
									response.write ""
								response.write "</td>"
							response.write "</tr>"
		
							If ubound(locarrAddReliefYears) > 0 then
								response.write "<tr>"
									response.write "<td>"
										response.write ""
									response.write "</td>"
									response.write "<td>"
										response.write "<br>"
									response.write "</td>"
									response.write "<td  colspan=3>"
										response.write "<select class='basicselect' name=fldcboAddYear>"
											loclngCounter = 0
											while loclngCounter < ubound(locarrAddReliefYears)
												loclngCounter = loclngCounter + 1
												response.write "<option value="""
													response.write locarrAddReliefYears(loclngCounter)
												response.write """>"
												response.write locarrAddReliefYears(loclngCounter)
												response.write "</option>"
											Wend
										response.write "</select>"
										response.write "&nbsp;<input type=submit name=btnSubmit value=""Add Relief"">"
									response.write "</td>"
									response.write "<td>"
										response.write ""
									response.write "</td>"
								response.write "</tr>"
							End If
							
							If ubound(locarrRemoveReliefYears) > 0 then
								response.write "<tr>"
									response.write "<td>"
										response.write ""
									response.write "</td>"
									response.write "<td>"
										response.write "<br>"
									response.write "</td>"
									response.write "<td colspan=3>"
										response.write "<select class='basicselect' name=fldcboRemoveYear>"
											loclngCounter = 0
											while loclngCounter < ubound(locarrRemoveReliefYears)
												loclngCounter = loclngCounter + 1
												response.write "<option value="""
													response.write locarrRemoveReliefYears(loclngCounter)
												response.write """>"
												response.write locarrRemoveReliefYears(loclngCounter)
												response.write "</option>"
											Wend
										response.write "</select>"
										response.write "&nbsp;<input type=submit name=btnSubmit value=""Remove Relief"">"
									response.write "</td>"
									response.write "<td>"
										response.write ""
									response.write "</td>"
								response.write "</tr>"
							End If
							
						response.write "</form>"
						response.write "<tr>"
							response.write "<td colspan=6>"
								response.write "<br>"
							response.write "</td>"
						response.write "</tr>"
			
					response.write "</table>"
					response.write ""
				End With
			Else
				mWriteELPEditForm locobjActiveELP
			End If
		End If
	end function
	

	'** WRITE ADMIN LEAVE PERIOD ***
	function mWriteAdminLeavePeriod(ByRef locobjUser, ByRef locobjLeaveRequest, ByVal locblnValidateRequest)
		Dim objColLeaveTypes
		Dim loclngCounter
		
		Set objColLeaveTypes = new cColLeaveTypes
		Set objColLeaveTypes.EE = locobjUser
		
		objColLeaveTypes.CollectionType = CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_ADMIN
		
		response.write "<form name=frmLeavePeriod action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"

			'**** START HIDDEN FIELD VALUES ****
			response.write "<input type=hidden name=formname value=frmLeavePeriod>"
			response.write "<input type=hidden name=ee value="""
				response.write locobjUser.WWID
			response.write """>"
			response.write "<input type=hidden name=m value="""
				response.write strMode
			response.write """>"
			response.write "<input type=hidden name=itemid value="""
				response.write locobjLeaveRequest.ID
			response.write """>"
			'**** END HIDDEN FIELD VALUES ****
		
			response.write "<br><table class='pageContentHeader'>"							
				response.write "<tr>"
					response.write "<td><h3>"
						if locobjLeaveRequest.ID = 0 then
							response.write "Add "
						else
							response.write "View "
						end if
						response.write "Leave Period"
					response.write "</h3></td>"
				response.write "</tr>"
			response.write "</table>"	

			
				If locblnValidateRequest then
					response.write "<table class='pageContentTable'>"	
					response.write "<tr>"
						response.write "<td>"
							If not locobjLeaveRequest.AdminLeavePeriodFormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeaveRequest.AdminLeavePeriodErrorMessage
							ElseIf not locobjLeaveRequest.NewAdminRequestIsValid then
								mWriteFormError CONST_FORM_ERROR_INVALID, locobjLeaveRequest.NewAdminRequestErrorMessage
							End If
						response.write "</td>"
					response.write "</tr>"
					response.write "</table>"	
				End If

			response.write "<table class='pageContentTable' style='width:450px'>"	
				response.write "<tr>"
					response.write "<td style='width:150px'>"
						response.write "<span class='th'>Leave Type</span>"
					response.write "</td>"
					response.write "<td style='width:300px'>"
						response.write "<select class='basicselect' name=fldstrLeaveType"
						if locobjLeaveRequest.ID <> 0 then
							response.write " disabled"
						end if
						response.write ">"
							loclngCounter = 0
							while loclngCounter < objColLeaveTypes.Count
								loclngCounter = loclngCounter + 1
								response.write "<option value='"
									response.write objColLeaveTypes.Item(loclngCounter).Name
									response.write "'"
									if objColLeaveTypes.Item(loclngCounter).Name = locobjLeaveRequest.LeaveType.Name then
										response.write " selected='selected'"
									end if
								response.write ">"
								response.write objColLeaveTypes.Item(loclngCounter).Name
							wend
						response.write "</select>"
					response.write "</td>"
				response.write "</tr>"	
			response.write "</table>"	
			
			response.write "<table class='pageContentTable' style='width:450px'>"
					response.write "<tr><td style='width:150px'>"
						response.write "<span class='th'>Start Date</span>"
					response.write "</td>"
					response.write "<td style='width:200px'>"
						response.write "<input placeholder='1 Jan 2001' class='evdatepicker' name=fldstrStartDate type=text size=11 maxlength=11 value="""' 
							response.write mFormatDate(locobjLeaveRequest.StartDate,"medium")
						response.write """"
						
						if locobjLeaveRequest.ID <> 0 then
							response.write " disabled"
						end if
						response.write ">"
					response.write "</td>"
					response.write "<td style='width:100px'>"
							mWriteTimeOption "fldstrStartTime", locobjLeaveRequest.StartTime, (locobjLeaveRequest.ID <> 0)
					response.write "</td>"
				response.write "</tr>"	
			response.write "</table>"	
			
			response.write "<table class='pageContentTable' style='width:450px'>"					
					response.write "<tr><td style='width:150px'>"
						response.write "<span class='th'>End Date</span>"
					response.write "</td>"
					response.write "<td style='width:200px'>"
						response.write "<input class='evdatepicker' placeholder='1 Jan 2001' name=fldstrEndDate type=text size=11 maxlength=11 value="""
							response.write mFormatDate(locobjLeaveRequest.EndDate,"medium")
						response.write """"
						if locobjLeaveRequest.ID <> 0 then
							response.write " disabled"
						end if
						response.write ">"		
					response.write "</td>"
					response.write "<td style='width:100px'>"
						mWriteTimeOption "fldstrEndTime", locobjLeaveRequest.EndTime, (locobjLeaveRequest.ID <> 0)
					response.write "</td>"
				response.write "</tr>"	
			response.write "</table>"	
			
			response.write "<table class='pageContentTable'>"				
				response.write "<tr>"
					response.write "<td>"
						response.write "<div class='th' style='margin-bottom:10px'>Comments</div>"
						response.write "<textarea name=fldstrComments "
						if locobjLeaveRequest.ID <> 0 then
							response.write " disabled"
						end if
						response.write ">"
							response.write mHTMLEncode(locobjLeaveRequest.RequestComments)
						response.write "</textarea><br>"
						
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<br>"
						if locobjLeaveRequest.ID = 0 then
							response.write "<input type=submit name=""btnSubmit"" value=""Save"">"
						else
							response.write "&nbsp;<input type=submit name=""btnSubmit"" onclick=""javascript:return(confirm('Are you sure you want to delete this leave request?'));"" value=""Delete"">"
						end if
						response.write "&nbsp;<input type=submit name=""btnSubmit"" onclick=""javascript:return(confirm('Are you sure you want to clear the form?'));"" value=""Cancel"">"
						response.write "<br>"
						response.write "<br>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</form>"
		response.write "</table>"
		response.write "<br><br>"
		
		Set objColLeaveTypes = nothing
		
	end function


	'*** WRITE ADMIN MENU ***	
	function mWriteAdminMenu(ByRef locobjSelectedUser, ByVal locstrCurrentSelection)
		With locobjSelectedUser
			response.write "<table class='pageContentWidth'>"
				response.write "<tr><td>"
				response.write "<div class='nav-menu'><ul>"
        
				response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" title=""Select the user you want to work with."">"
				response.write "Select User"
				response.write "</a></li>"
				
				if request("ee") <> "" Then

					response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?ee="
						response.write .WWID
					response.write """ title=""Edit the selected user's profile."">"
					response.write "User Profile"
					response.write "</a></li>"						
		
					response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=co&amp;ee="
						response.write .WWID
					response.write """ title=""Edit the selected user's Carry Over leave."">"
					response.write "Carry Over/Exceptions"
					response.write "</a></li>"		
		
					response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=lv&amp;ee="
							response.write .WWID
					response.write """ title=""Edit the selected user's leave periods."">"
					response.write "Leave"
					response.write "</a></li>"						
						
					response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=elp&amp;ee="
						response.write .WWID
					response.write """ title=""Update the selected user's ELP details."">"
					response.write "ELP"
					response.write "</a></li>"
				
				End If
				
	
				response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=rp"" title=""Run reports for e-Vacation."">"
					response.write "Reports"
				response.write "</a></li>"
										
										
				response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=dq"" title=""Send a query to the database"">"
					response.write "Database Query"
				response.write "</a></li>"
					
				
				response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=sfm""  title=""Update SQL Tables Through Forms."">"
					response.write "SQL Forms"
				response.write "</a></li>"

				
				response.write "<li style='margin-left:0px'><a class='iframe' href='" & CONST_APPLICATION_PATH & "/adminguide.htm' title=""Click here for e-Vacation's Admin Guide."">"
					response.write "Admin Guide"
				response.write "</a></li>"
				response.write "</ul></div></td></tr>"
			response.write "</table>"
		End With
		response.write ""
	end function
	

	'** WRITE ADMIN SELECT USER FORM ***
	function mWriteAdminSelectUserForm(ByRef locobjUser, ByVal locblnValidateUser)
		response.write "<form name=frmSelectUser action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"
		
		'**** START HIDDEN FIELD VALUES ****
		response.write "<input type=hidden name=formname value=frmSelectUser>"
		'**** END HIDDEN FIELD VALUES ****
		
		response.write "<br><table class='pageContentHeader'><tr><td><h3>Select User</h3></td></tr></table>"			
	
		response.write "<table class='pageContentTable' style='width:800px'>"		
			If locblnValidateUser then
				response.write "<tr>"
					response.write "<td colspan='3'>"
						If not locobjUser.AdminSelectUserFormIsValid then
							mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjUser.AdminSelectUserFormErrorMessage
						End If
					response.write "</td>"
				response.write "</tr>"
			End If

			response.write "<tr>"
				response.write "<td style='width:150px'>"
					response.write "<span class='th'>Select User</span>"
				response.write "</td><td class='width:500px'>"
				
					Dim eloclngCounter
					Dim eloclngResults
					Dim EmployeesAll
					Set EmployeesAll = new cColEmployeesAll				
					
					response.write "  <div class='ui-widget'>"
					response.write "  <select class='wwidcombobox' name=ee>"
					response.write "	<option value=''>Select one...</option>"
					eloclngCounter = 0
					eloclngResults = EmployeesAll.Count
					While eloclngCounter < eloclngResults
						eloclngCounter = eloclngCounter + 1
						with EmployeesAll.Item(eloclngCounter)											
							response.write "<option value='" & .WWID & "' "
							response.write ">" & .LastNm & ", " & .FirstNm & " (" & .WWID & ")</option>"
						end with
					Wend
					response.write "</select></div>"
				response.write "</td>"	
				response.write "<td class='text-align:left;width:150px'>"	
					response.write "<input type=submit value=""Select"">"
				response.write "</td>"
				response.write "</tr>"
			response.write "</table><br>"						
		response.write "</form>"
    
    response.write "<form id='addwwidform' name=frmSelectUser action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"
    response.write "<table class='pageContentHeader'><tr><td><h3>Add User</h3></td></tr></table>"	   
    
    response.write "<table class='pageContentTable' style='width:600px'>"	

			response.write "<tr>"
				response.write "<td style='width:150px'>"
					response.write "<span class='th'>User WWID</span>"
				response.write "</td><td class='width:300px'><input id='userwwid' type='text' value='' name='ee'/>"
				response.write "</td>"	
				response.write "<td class='text-align:left;width:150px'>"	
					response.write "<input type=button id='submitwwid' value=""Add User"">"
				response.write "</td>"
				response.write "</tr>"
      response.write "<tr>"
				response.write "<td style='width:150px'>"
				response.write "</td><td class='width:300px'>"
				response.write " <span id='addwwidfeedback' style='color:#AA0000;'></span></td>"	
				response.write "<td class='text-align:left;width:150px'>"	
				response.write "</td>"
				response.write "</tr>"
			response.write "</table><br><br>"						
		response.write "</form>"
		
	end function



	'*** WRITE FORM to add bank holidays ***'[MFILLAST 08-2006] bh TODO
	function mWriteBankHolidayForm(ByVal locblnValidateDate)
	
			response.write "<table class='pageContentHeader'>"				
				response.write "<tr>"
					response.write "<td>"
						response.write "<h3>Add a bank holiday</h3>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"

					If locblnValidateDate then
						response.write "<table class='pageContentHeader'>"
						response.write "<tr>"
							response.write "<td colspan=2>"
								If not locobjUser.AdminUserProfileFormIsValid then'MODIF HERE to check date + look at the call
									mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjUser.AdminUserProfileFormErrorMessage
								End If
							response.write "</td>"
						response.write "</tr>"
						response.write "</table>"
					End If
			
				response.write "<form name=""frmBankHoliday"" method=post action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"">"
					response.write "<input type=hidden name=formname value=""frmBankHoliday"">"
					response.write "<input type=hidden name=m value="""
						response.write strMode
					response.write """>"
				
				response.write "<table class='pageContentTable'>"
					response.write "<tr>"
						response.write "<td style='text-align:right'>"
							response.write "Date"
						response.write "</td>"
						response.write "<td align=left>&nbsp;"
							response.write "<input class='evdatepicker' name=fldDateBankHoliday type=text size=11 maxlength=11 value=''>"							
						response.write "</td>"
					response.write "</tr>"				
					response.write "<tr>"
						response.write "<td colspan=2>"
							response.write "<input type=submit value=""Add date""><br><br>"
						response.write "</td>"
					response.write "</tr>"
				response.write "</form>"
			response.write "</table>"
		response.write "<br><br>"
	end function
	'end bank holidays
	
	
	'*** WRITE FORM to send querys to database ***'[MFILLAST 08-2006]
	function mWriteDatabaseQueryForm()
	
		response.write "<form name=""frmDatabaseQuery"" method=post action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"">"
		response.write "<input type=hidden name=formname value=""frmDatabaseQuery"">"
		response.write "<input type=hidden name=m value=""dq"">"
		
			'display the form
			response.write "<br><table class='pageContentHeader'>"				
					response.write "<tr>"
						response.write "<td>"
							response.write "<h3>SQL Query</h3>"
						response.write "</td>"
					response.write "</tr>"
			response.write "</table>"

			response.write "<table class='pageContentTable'>"
					response.write "<tr>"
						response.write "<td>"
						response.write "<div class='th' style='margin-bottom:10px'>Query</div>"
						response.write "<textarea name=""fldstrQuery"">"
						response.write "</textarea><br>"
						response.write "Be careful with the query : no test is done before executing"
						response.write "<br><br>"
						response.write "<a class='iframe' href='" & CONST_APPLICATION_PATH & "/query.htm' title=""Examples of queries"">See examples of queries</a>"
						response.write ""
					response.write "</td>"
					
				response.write "</tr>"

					response.write "<tr>"
						response.write "<td>"
							response.write "<input type=button onclick=""javascript:executeQuery('frmDatabaseQuery','fldstrQuery');"" value=""Execute"">"
						response.write "</td>"
					response.write "</tr>"
				response.write "</form>"
			response.write "</table>"
		response.write "<br><br>"
	end function
'end database update
	
	
	
	
	'*** WRITE ADMIN USER PROFILE FORM ***
	function mWriteAdminUserProfileForm(ByRef locobjUser, ByVal locblnValidateProfile)
		With locobjUser
				response.write "<table class='pageContentHeader'>"
				mDebugPrint "IS VALID USER='" & locobjUser.IsValidUser & "'<br>"
				response.write "<form name=""frmUserProfile"" method=post action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"">"
					response.write "<input type=hidden name=formname value=""frmUserProfile"">"
					response.write "<input type=hidden name=m value="""
						response.write strMode
					response.write """>"
					response.write "<input type=hidden name=ee value="""
						response.write .WWID
					response.write """>"
					response.write "<tr>"
						response.write "<td><div class='th'><br>"
							response.write "<div span='th'>User Profile "
							response.write .FullName
						response.write "</div></td>"
					response.write "</tr>"
				response.write "</table>"
				
				If locblnValidateProfile then
					response.write "<table class='pageContentTable'>"
					response.write "<tr>"
						response.write "<td>"
							If not locobjUser.AdminUserProfileFormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjUser.AdminUserProfileFormErrorMessage
							End If
						response.write "</td>"
					response.write "</tr>"
					response.write "</table>"
				End If
	
				response.write "<table class='pageContentTable' style='width:500px'>"
					response.write "<tr>"
						response.write "<td>"
							response.write "<span class='th'>End of contract</span>"
						response.write "</td>"
						response.write "<td>"
						response.write "<input id='dpec' class='evdatepicker' name='fldendDate' type='text' size='11' maxlength='11' value='"
								response.write mFormatDate(.endDate,"datepicker")
							response.write "'>"
            if IsDate(.endDate) then 
              response.write "<script type='text/javascript'>"
              response.write "$(document).ready( function(){"
                response.write "$( '#dpec' ).datepicker('setDate', new Date("
                  response.write mFormatDate(.endDate,"datepicker")
                response.write "));"
              response.write "});</script>"
            End If              
						response.write "</td>"
					response.write "</tr>"
					response.write "<tr>"
						response.write "<td>"
							response.write "<span class='th'>Leave Tracking Enabled</span>"
						response.write "</td>"
						response.write "<td align=left>"
							response.write "<input name=fldblnIsEELeaveTracked value=""True"" type=checkbox"
								if .IsEELeaveTracked then
									response.write " checked"
								end if
							response.write ">"
						response.write "</td>"
					response.write "</tr>"
					response.write "<tr>"
						response.write "<td>"
							response.write "<span class='th'>Administrator Access</span>"
						response.write "</td>"
						response.write "<td>"
							response.write "<input name=fldblnIsAdmin value=""True"" type=checkbox"
								if .IsAdmin then
									response.write " checked"
								end if
							response.write ">"
						response.write "</td>"
					response.write "</tr>"
					response.write "<tr>"
						response.write "<td><br>"
							response.write "<input type=submit value=""Save Changes""><br><br>"
						response.write "</td>"
					response.write "</tr>"
				response.write "</form>"
			response.write "</table>"
		End With
		response.write "<br><br>"
	end function
	

	'*** WRITE APPROVE LEAVE REQUEST ***
	function mWriteApproveLeaveRequest(ByRef locobjLeaveRequest, ByVal locblnValidateApproval)
		' [MOF 24/11/08] Added link to Javascript
			
			With locobjLeaveRequest
				response.write "<form name=frmRequestLeave action=""" & CONST_APPLICATION_PATH & "/approverequests.asp"" method=post>"
					'**** START HIDDEN FIELD VALUES ****
					response.write "<input type=hidden name=formname value=frmLeaveApproval>"
					response.write "<input type=hidden name=m value=""approverequest"">"
					response.write "<input type=hidden name=itemid value="""
						response.write .ID
					response.write """>"
					'**** END HIDDEN FIELD VALUES ****
					
					response.write "<table class='pageContentHeader'><tr><td><h3>"
					
					If  .Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED then
						response.write "Approve Cancellation Request"
					Else
						response.write "Approve Leave Request"
					End If
					
					response.write "</h3></td></tr></table>"
					
					response.write "<table class='pageContentTable'><tr><td>"

					If locblnValidateApproval then
						response.write "<tr>"
							response.write "<td style='text-align:center' colspan=9>"
								If not locobjLeaveRequest.ApprovalFormIsValid then
									mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeaveRequest.ApprovalErrorMessage
								End If
							response.write "</td>"
						response.write "</tr>"
					End If
					
					response.write "<tr>"
						response.write "<td colspan=2>"
							response.write "<span class='th'>Employee </span> "
						response.write "</td>"
						response.write "<td colspan=7>"
							response.write .EE.FullName
							response.write " (WWID: "
							response.write .EE.WWID
							response.write ")"
						response.write "</td>"
					response.write "</tr>"
					
					response.write "<tr>"
						response.write "<td colspan=2>"
							response.write "<span class='th'>Leave Type</span> "
						response.write "</td>"
						response.write "<td>"
							response.write .LeaveType.Name
						response.write "</td>"
						response.write "<td>"
							response.write "<span class='th'>Start Date</span> "
						response.write "</td>"
						response.write "<td>"
							response.write mFormatDate(.StartDate,"medium with day")
							response.write " "
							response.write .StartTime
						response.write "</td>"
						response.write "<td>"
							response.write "<span class='th'>End Date</span> "
						response.write "</td>"
						response.write "<td>"
							response.write mFormatDate(.EndDate,"medium with day")
							response.write " "
							response.write .EndTime
						response.write "</td>"
						response.write "<td>"
							response.write "<span class='th'>Days</span> "
						response.write "</td>"
						response.write "<td>"
							response.write .Days
						response.write "</td>"
					response.write "</tr>"
					
					response.write "<tr>"
						response.write "<td colspan=9>"
							response.write "<div class='th' style='margin-bottom:10px'>Employee's Comments</div>"
							response.write mHTMLEncode(.RequestComments)
						response.write "<br><br></td>"
					response.write "</tr>"

					response.write "<tr>"
						response.write "<td colspan=9>"
							response.write "<div class='th'>Comments to be included in your response</div><br>"
						response.write "<textarea name=fldstrComments >"
							response.write mHTMLEncode(locobjLeaveRequest.ResponseComments)
						response.write "</textarea><br>"						
						response.write "</td>"
					response.write "</tr>"
					
					response.write "<tr>"
						response.write "<td colspan=9>"
							response.write "<input id='approveButton' name=btnSubmit type=submit value=""Approve"">"
							response.write " "
							response.write "<input id='rejectButton' name=btnSubmit type=submit value=""Reject"">"
						response.write "</td>"
					response.write "</tr>"
				response.write "</form>"
			End With
		response.write "</table><br><br>"
		response.write ""
	end function


	'**** WRITE APPROVE REQUESTS ****
	function mWriteApproveRequests(ByRef locobjUser)
		Dim loclngCounter
		Dim locstrClassText
		
		response.write "<br><table class='pageContentHeader'><tr><td><h3>Approve Requests</h3></td></tr></table>"
		
		response.write "<table class='pageContentTable'>"
			
			With locobjUser.Approvals
				If .Count = 0 then
					response.write "<tr>"
						response.write "<td>"
							response.write "<div class='feedbackbox'>You currently have no requests pending your approval.</div>"
						response.write "</td>"
					response.write "</tr>"
				Else
					response.write "<tr>"
						response.write "<th>"
							response.write "<br>"
						response.write "</th>"
						response.write "<th>"
							response.write "Employee"
						response.write "</th>"
						response.write "<th>"
							response.write "Status"
						response.write "</th>"
						response.write "<th>"
							response.write "Start Date"
						response.write "</th>"
						response.write "<th>"
							response.write "End Date"
						response.write "</th>"
						response.write "<th>"
							response.write "Days"
						response.write "</th>"
						response.write "<th>"
							response.write "Appointed Approver"
						response.write "</th>"
						response.write "<th>"
							response.write "<br>"
						response.write "</th>"
						response.write "<th>"
							response.write "<br>"
						response.write "</th>"
					response.write "</tr>"
				
					loclngCounter = 0
					While loclngCounter < .Count
						loclngCounter = loclngCounter + 1
						If .Item(loclngCounter).AppointedApprover.WWID = objCurrentUser.WWID then
							locstrClassText = ""
						else
							locstrClassText = " class=txtdim"
						end if
						response.write "<tr>"
							response.write "<td>"
								response.write "<br>"
							response.write "</td>"
							response.write "<td "
								response.write locstrClassText
							response.write ">"
								response.write .Item(loclngCounter).EE.FullName
							response.write "</td>"
							response.write "<td "
								response.write locstrClassText
							response.write ">"
								response.write .Item(loclngCounter).Status
							response.write "</td>"
							response.write "<td "
								response.write locstrClassText
							response.write ">"
								response.write mFormatDate(.Item(loclngCounter).StartDate,"medium")
								response.write "&nbsp;"
								response.write .Item(loclngCounter).StartTime
							response.write "</td>"
							response.write "<td "
								response.write locstrClassText
							response.write ">"
								response.write mFormatDate(.Item(loclngCounter).EndDate,"medium")
								response.write "&nbsp;"
								response.write .Item(loclngCounter).EndTime
							response.write "</td>"
							response.write "<td "
								response.write locstrClassText
							response.write ">"
								response.write FormatNumber(.Item(loclngCounter).Days,1)
							response.write "</td>"
							response.write "<td "
								response.write locstrClassText
							response.write ">"
								response.write .Item(loclngCounter).AppointedApprover.FullName
							response.write "</td>"
							response.write "<td "
								response.write locstrClassText
							response.write ">"
								response.write "<a href=""" & CONST_APPLICATION_PATH & "/approverequests.asp?m="
								response.write CONST_MODE_APPROVE_REQUESTS_VIEW_REQUEST
								response.write "&itemid="
								response.write .Item(loclngCounter).ID
								response.write """ title=""Click here to view more details about this request and to approve or reject it."">View</a>"
							response.write "</td>"
							response.write "<td>"
								response.write "<br>"
							response.write "</td>"
						response.write "</tr>"
					Wend
					response.write "</table>"
					response.write ""
					response.write "<table class='pageContentTable'>"
						response.write "<tr>"
							response.write "<td>"
								response.write "<br>"
								response.write "<b>Key:</b>&nbsp;&nbsp;&nbsp;&nbsp;<span class=txtdim>You are not the appointed approver for this request.</span>&nbsp;&nbsp;&nbsp;&nbsp;You are the appointed approver for this request.<br>"
								response.write "<br>"
							response.write "</td>"
						response.write "</tr>"
				end if
			End With
		response.write "</table>"
		response.write ""
	end function


	'** WRITE CANCEL LEAVE PERIOD REQUEST FORM ***
	function mWriteCancelLeavePeriodRequestForm(ByRef locobjUser, ByRef locobjLeavePeriod, ByVal locblnValidateRequest)
	
		
			response.write "<form name=frmCancelLeavePeriod action=""" & CONST_APPLICATION_PATH & "/leavesummary.asp"" method=post>"

				'**** START HIDDEN FIELD VALUES ****
				response.write "<input type=hidden name=formname value=frmCancelLeavePeriod>"
				response.write "<input type=hidden name=ee value="""
					response.write objEEtoView.WWID
				response.write """>"
				response.write "<input type=hidden name=m value="""
					response.write strMode
				response.write """>"

				response.write "<input type=hidden name=itemid value="""
					response.write locobjLeavePeriod.ID
				response.write """>"
				
				'**** END HIDDEN FIELD VALUES ****
				
			response.write "<table class='pageContentHeader'>"				
				response.write "<tr>"
					response.write "<td><h3>"
						response.write "Cancel Leave Period</h3>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"
			
					
				If locblnValidateRequest then
					response.write "<table class='pageContentTable'>"	
					response.write "<tr>"
						response.write "<td>"
							If not locobjLeavePeriod.CancelRequestFormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeavePeriod.CancelRequestFormErrorMessage
							End If
						response.write "</td>"
					response.write "</tr>"
					response.write "</table>"
				End If
			
			response.write "<table class='pageContentTable'>"	
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Leave Type</span> "
						response.write locobjLeavePeriod.LeaveType.Name
					response.write "</td>"
					response.write "<td style='text-align:right'>"
						response.write "<span class='th'>Start Date</span>"
					response.write "</td>"
					response.write "<td align=left>"
							response.write mFormatDate(locobjLeavePeriod.StartDate,"medium with day")
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<div class='th' style='margin-bottom:10px'>Alternative Approver</div>"
						response.write "<input class='evdatepicker' name=fldstrApproverWWID type=text size=8 maxlength=8 value="""
							response.write locobjLeavePeriod.Approver.WWID
						response.write """>"
					response.write "</td>"
					response.write "<td style='text-align:right'>"
						response.write "<span class='th'>End Date</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(locobjLeavePeriod.EndDate,"medium with day")
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td colspan=3>"
						response.write "<span class='th'>Comments</span><br>"
						response.write "<textarea name=fldstrComments>"
							response.write mHTMLEncode(locobjLeavePeriod.RequestComments)
						response.write "</textarea><br>"
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td colspan=3>"
						response.write "<input type=checkbox "
							if locobjLeavePeriod.EmailConfOfRequestReq then
								response.write "checked "
							end if
						response.write "name=fldblnNotify value=True>"
						response.write "&nbsp;Send me an e-mail confirmation of this cancellation request."
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td colspan=3>"
						response.write "<br>"
						response.write "<input type=submit value=""Submit Cancellation Request""><br>"
						response.write "<br>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</form>"
		response.write "</table>"
		response.write "<br><br>"
		
	end function

	'**** WRITE DELEGATE FORM ****
	function mWriteDelegateForm(ByRef locobjUser, ByVal locblnValidate)
	
		
		Dim locstrFormName

		If locobjUser.HasDelegate then
			If locblnValidate then
				locstrFormName = "frmAppointDelegate"
			Else
				locstrFormName = "frmRemoveDelegate"
			End If
		else
			locstrFormName = "frmAppointDelegate"
		end if

		response.write "<table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td><h3>"
					response.write "Delegate Form"
				response.write "</h3></td>"
			response.write "</tr>"
		response.write "</table>"
		
		response.write "<table class='pageContentTable'>"
			response.write "<form name="
			response.write locstrFormName
			response.write " action=""" & CONST_APPLICATION_PATH & "/delegate.asp"" method=post>"

				'**** START HIDDEN FIELD VALUES ****
				response.write "<input type=hidden name=formname value="
				response.write locstrFormName
				response.write ">"
				response.write "<input type=hidden name=ee value="""
					response.write locobjUser.WWID
				response.write """>"
				'**** END HIDDEN FIELD VALUES ****
	
				If locblnValidate then
					response.write "<tr>"
						response.write "<td><div class='feedbackerror'>"
							mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjUser.DelegateFormErrorMessage
						response.write "</div></td>"
					response.write "</tr>"
				End If
	
				'**** DISPLAY FORM FOR REMOVING DELEGATE
				If locstrFormName = "frmRemoveDelegate" then
 					response.write "<tr>"
						response.write "<td>"
							response.write "<br><span class='th'>Active Delegate:</span> "
							response.write locobjUser.ActiveDelegate.FullName
							response.write " ("
							response.write locobjUser.ActiveDelegate.WWID
							response.write ")"
							response.write "<br>"
							response.write "<br>"
						response.write "</td>"
					response.write "</tr>"	
					response.write "<tr>"
						response.write "<td>"
							response.write "<input type=submit value=""Remove Delegate""><br>"
							response.write "<br>"
						response.write "</td>"
					response.write "</tr>"	
					
				'**** DISPLAY FORM FOR APPOINTING DELEGATE
				Else
									
					response.write "<tr>"
						response.write "<td>"						
							
							Dim eloclngCounter
							Dim eloclngResults
							Dim EmployeesAll
							Set EmployeesAll = new cColEmployeesAll				
							
							response.write "<div class='th' style='margin-bottom:10px'>Choose a delegate</div>"
							response.write "  <div class='ui-widget' style='width:100%'>"
							response.write "  <select class='wwidcombobox' name=fldstrDelegateWWID>"
							response.write "	<option value=''>Select one...</option>"
							eloclngCounter = 0
							eloclngResults = EmployeesAll.Count
							While eloclngCounter < eloclngResults
								eloclngCounter = eloclngCounter + 1
								with EmployeesAll.Item(eloclngCounter)											
									response.write "<option value='" & .WWID & "' "
									response.write ">" & .LastNm & ", " & .FirstNm & " (" & .WWID & ")</option>"
								end with
							Wend
							response.write "</select></div>"													
							
						response.write "</td>"
					response.write "</tr>"
					
					response.write "<tr>"
						response.write "<td>"
							response.write "<input type=submit value=""Appoint Delegate"">"
						response.write "</td>"
					response.write "</tr>"	
					
				End If
			response.write "</form>"
		response.write "</table><br><br>"
	end function
	

	'*** WRITE ELP DETAILS ***	
	function mWriteELPDetails(ByRef locobjELP, ByVal locblnAllowEdit)
		Dim loclngYear
		Dim loclngEndYear
		With locobjELP
			
			response.write "<table class='pageContentHeader'>"
				response.write "<tr>"
					response.write "<td><h3>"
						response.write .Status
						response.write " ELP"
					response.write "</h3></td>"
				response.write "</tr>"
			response.write "</table>"
			
			response.write "<table class='pageContentTable'>"
			
				response.write "<tr>"
					response.write "<td style='width:200px'>"
						response.write "<span class='th'>Activated On<span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(.DateActivated,"medium with day")
					response.write "</td>"
				response.write "</tr>"
				
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Activated by</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .ActivatedBy.FullName
						response.write " ("
						response.write .ActivatedBy.WWID
						response.write ")"
					response.write "</td>"
				response.write "</tr>"
	
				loclngYear = DatePart("yyyy",.DateActivated)
				loclngEndYear = DatePart("yyyy",.MaturityDate)
				While loclngYear <= loclngEndYear
									
					If loclngYear = DatePart("yyyy",.DateActivated) then
						response.write "<tr>"
							response.write "<td>"
								response.write "<span class='th'>Days Banked</span>"
							response.write "</td>"
						response.write "</tr>"
					End If
					
					response.write "<tr>"
						If loclngYear = DatePart("yyyy",.DateActivated) then
							response.write "<td>"
								response.write "<span class='th' style='padding-left:15px'>On Activation</span>"
							response.write "</td>"
							response.write "<td>"
								response.write .DaysBankedOnActivation
							response.write "</td>"
						ElseIf .ReliefInYear(loclngYear) then
							response.write "<td>"
								response.write "<span class='th' style='padding-left:15px'>In "
								response.write loclngYear
								response.write "</span>"
							response.write "</td>"
							response.write "<td>"
								response.write "0 (relief given)"
							response.write "</td>"
						Else
							response.write "<td>"
								response.write "<span class='th' style='padding-left:15px'>In "
								response.write loclngYear
								response.write "</span>"
							response.write "</td>"
							response.write "<td>"
								response.write .DaysBankedPerYear
								If loclngYear > DatePart("yyyy", Date()) then
									response.write " (expected)"
								End If
							response.write "</td>"
						End If
						
					response.write "</tr>"
					loclngYear = loclngYear + 1
				Wend

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Total Days To Bank</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .TargetDays
					response.write "</td>"
				response.write "</tr>"

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Mature"
							If .MaturityDate <= date() then
								response.write "d"
							Else
								response.write "s"
							End If
						response.write "</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(.MaturityDate,"medium")
					response.write "</td>"
				response.write "</tr>"
				
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Expire"
							If .ExpiryDate <= date() then
								response.write "d"
							Else
								response.write "s"
							End If
						response.write "</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(.ExpiryDate,"medium")
					response.write "</td>"
				response.write "</tr>"

				If locblnAllowEdit and .Status = "Active" then
					response.write "<tr>"
						response.write "<form method=get action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"">"
							response.write "<td colspan='100%'>"
								response.write "<input type=hidden name=m value=""eelp"">"
								response.write "<input type=hidden name=ee value="""
									response.write objEEtoView.WWID
								response.write """>"
								response.write "<input type=hidden name=itemid value="""
									response.write .ELPID
								response.write """>"
								response.write "<input type=submit value=""Edit ELP Activation Details"" title=""Edit the activation details for this ELP instance."">"
						response.write "</td>"
						response.write "</form>"
					response.write "</tr>"
				End If
					
			response.write "</table>"
			response.write "<br><br>"
		end with
	end function


	'*** WRITE ELP EDIT FORM ***
	function mWriteELPEditForm(ByVal locobjELPInstance)
	
			With locobjELPInstance
				If locblnValidateActivateELP then
					response.write "<table class='pageContentHeader'>"
						response.write "<tr>"
							response.write "<td>"
								response.write "<br>"
							response.write "</td>"
							response.write "<td style='text-align:center' colspan=3>"
								If not .AdminActivateELPFormIsValid then
									mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, .AdminActivateELPFormErrorMessage
								End If
							response.write "</td>"
						response.write "</tr>"
					response.write "</table>"
				End If

				response.write "<form name=frmActivateELP action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"
					'**** START HIDDEN FIELD VALUES ****
					response.write "<input type=hidden name=formname value=frmActivateELP>"
					response.write "<input type=hidden name=ee value="""
						response.write locobjUser.WWID
					response.write """>"
					response.write "<input type=hidden name=m value="""
						response.write strMode
					response.write """>"
					response.write "<input type=hidden name=itemid value="""
						response.write .ELPID
					response.write """>"
					'**** END HIDDEN FIELD VALUES ****
				
				response.write "<table class='pageContentHeader'>"
					response.write "<tr>"
						response.write "<td><h3>"
							'Response.Write ("****ELPID: " & .ELPID)
							If mGetSafeLongInteger(.ELPID,0) <> 0 then
								response.write "Edit"
							else
								response.write "Activate"
							end if
							response.write " ELP"
						response.write "</h3></td>"
					response.write "</tr>"
				response.write "</table>"
				
				response.write "<table class='pageContentTable'>"	
				
					response.write "<tr>"	
						response.write "<td>"						
							response.write "<span class='th'>Activation Date</span> "
							response.write " <input class='evdatepicker' name=fldstrActivationDate type=text size=11 maxlength=11 value="""
								response.write mFormatDate(.DateActivated,"medium")
							response.write """>"							
						response.write "</td>"					
					response.write "</tr>"
					
					response.write "<tr>"						
						response.write "<td>"
							response.write "<span class='th'>Bank</span> "
								response.write " <input name=flddblDaysBankedOnActivation type=text size=2 maxlength=2 value="""
									response.write .DaysBankedOnActivation
								response.write """>"
							response.write " <span class='th'>days on activation</span>"
						response.write "</td>"
					response.write "</tr>"

					response.write "<tr>"
						response.write "<td colspan=2>"
							response.write "<span class='th'>Total Days To Be Banked</span> "
							response.write " <input name=flddblTargetDays type=text size=2 maxlength=2 value="""
								response.write .TargetDays
							response.write """>"
						response.write "</td>"
					response.write "</tr>"
					
					response.write "<tr>"
						response.write "<td>"
							response.write "<input type=submit value="""

							If mGetSafeLongInteger(.ELPID,0) <> 0 then
								'response.write "Edit"
								Response.Write "Save"
							else
								response.write "Activate"
							end if
							
							response.write " ELP"" onclick=""javascript:return(confirm('Are you sure you want to "

							If mGetSafeLongInteger(.ELPID,0) <> 0 then
								response.write "edit this"
							else
								response.write "activate"
							end if
							
							response.write " ELP for this employee with the information entered?'))"">"
						response.write "</td>"
					response.write "</tr>"
			End With

			response.write "</table>"		
		response.write "</form>"
		response.write "<br><br>"

	end function


	'**** WRITE FORM ERROR ****
	function mWriteFormError(ByVal loclngErrorType, ByVal locstrMessage)
		Dim locstrHeading
		Select Case loclngErrorType
			Case CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY
				locstrHeading = "The form is not completed correctly" 
			Case CONST_FORM_ERROR_INVALID
				locstrHeading = "The information entered is invalid"
			Case Else
				locstrHeading = "An error occurred"
		End Select
		response.write "<div class='feedbackboxerror'><b>The form is not completed correctly - click <a style='cursor:pointer' id='errorshowlink'>here</a> for details</b>.</div>"
		
		locstrMessage = Replace(locstrMessage, "\n", "<br/>")
		response.write "<div id='dialog-message' title='" & locstrHeading &"'>"
		response.write "<p>"
		response.write locstrMessage
		response.write "</p>"
		response.write "</div>"
		
	end function
	
	
	'**** WRITE GENERAL ERROR ****
	function mWriteGeneralError(ByVal locstrMessage, ByVal locblnBackButtonMessage)
		response.write "<div class='feedbackbox'>"
			response.write locstrMessage
			if locblnBackButtonMessage then
				response.write "<br><br><br>"
				response.write "To try again, please press the 'Back' button on your browser."
			else
				response.write "<br><br>"
				response.write "Please try again."
			end if
		response.write "</div>"
	end function
	
	
	'**** WRITE USER INFO ****
	function mWriteUserInfo()
		Dim loclngCounter
		With objCurrentUser
	
             response.write "<form name=frmShareWithTeamCalendar action=""" & CONST_APPLICATION_PATH & "/userinfo.asp"" method=post>"	   

			response.write "<br><table class='pageContentHeader'><tbody><tr><td><h3>It's Me!</h3></td></tr></tbody></table>"
      
			response.write "<table class='pageContentTable'><tr><td style='vertical-align:top'><img style='height:200px;margin-left:auto;margin-right:auto;margin-top:30px;' src='common/phpldap/userimage.php?user=" & objCurrentUser.IDSID & "' alt=''></td><td>"
			response.write "<table>"
				response.write "<tr>"
					response.write "<td style='width:200px' colspan='2'>"
						response.write "<h4>User Information</h4>"				
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td style='width:200px'>"
						response.write "<span class='th'>WWID</span>"
					response.write "</td>"
					response.write "<td style='width:300px'>"
						response.write .WWID
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>NT Logon ID</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .IDSID
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>E-mail Address</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .Email
					response.write "</td>"
				response.write "</tr>"
        response.write "<tr><td><br></td></tr>"
				response.write "<tr>"
					response.write "<td style='width:200px' colspan='2'>"
						response.write "<h4>Personal Details</h4>"				
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td style='width:200px'>"
						response.write "<span class='th'>End of contract</span>"
					response.write "</td>"
					response.write "<td style='width:300px'>"
						if not isdate(objCurrentUser.endDate) then
								response.write "<i>not specified</i>"
						else
							response.write mFormatDate(objCurrentUser.endDate,"medium")
						end if
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Manager</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .Manager.FullName
						response.write " (WWID: "
						response.write .Manager.WWID
						response.write ")"
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Intel Entity Code</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .CompanyCd
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Blue Badge</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mIf(.IsBlueBadge, "Yes", "No")
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Part Time</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mIf(.IsPartTimer, "Yes", "No")
					response.write "</td>"
				response.write "</tr>"  
        response.write "<tr><td><br></td></tr>"

        response.write "<tr>"
					response.write "<td style='width:200px'>"
						response.write "<span class='th'>Delegate For</span>"				
					response.write "</td>"
					
					If not .IsDelegate then
							response.write "<td style='width:450px'>"
								response.write "(you are not currently appointed as a delegate for any managers on e-Vacation)"
							response.write "</td>"
							response.write "<td style='width:50px'>"
								response.write ""
							response.write "</td>"
						response.write "</tr>"
					Else
						loclngCounter = 0
						while loclngCounter < .DelegateForManagers.Count
							loclngCounter = loclngCounter + 1
								response.write "<td style='wdith:450px'>"
									response.write .DelegateForManagers.Item(loclngCounter).FullName
								response.write "</td>"
								if loclngCounter = 1 then
									response.write "<td style='width:50px'>"
										response.write ""
									response.write "</td>"
								end if
							response.write "</tr>"
						wend
					End If
     response.write "<tr><td><br></td></tr>"

        response.write "<tr>"
			response.write "<td style='width:200px' colspan='2'>"
				response.write "<h4>Privacy</h4>"				
			response.write "</td>"
		response.write "</tr>"

        response.write "<tr>"
                response.write "<td>"
					response.write "<span class='th'>Share Leave in Team Calendar</span>"	
				response.write "</td>"

                response.write "<td>"    
                    response.write "<input type=hidden name=formname value=frmShareWithTeamCalendar>"
                    response.write "<input type=hidden name=wwid value="""
                    response.write .WWID
                    response.write """>"   

                    Dim locCmd
                    Dim locRS        

                    Set locCmd = Server.CreateObject("ADODB.Command")
                    Set locCmd.ActiveConnection = glbConnection
                    locCmd.CommandType = adCmdStoredProc

                    locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, .WWID)
                    locCmd.CommandText = "usp_get_share_with_team_calendar"
                    Set locRS = locCmd.Execute               

		            response.write "<input type='checkbox' name=fldblnShare "
                        if locRs("shareWithTeamCalendar") then
				            response.write "checked "
			            end if
                    response.write ">"

                    locRS.close 

                response.write "</td>"
        response.write "</tr>"
     response.write "<tr>"
                response.write "<td><br>"
					
				response.write "</td>"

                response.write "<td>"
                     response.write "<input type=submit value=""Save"">"
                response.write"</td>"
        response.write "</tr>"

					response.write "</table>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
     response.write "</form>"
    
		End With
		
	end function
	

	function mWriteHMTLTop(ByVal locstrTitle)
	
		response.write "<!DOCTYPE html>" & vbCrLf	
		response.write "<html xmlns='http://www.w3.org/1999/xhtml'>" & vbCrLf	
		response.write "<head>"		 & vbCrLf	
		response.write "<title>" 
		response.write "e-Vacation - " 
		response.write locstrTitle 
		response.write "</title>" & vbCrLf	
		response.write "<meta name='viewport' content='width=device-width, initial-scale=1'>" & vbCrLf	
		response.write "<meta http-equiv='X-UA-Compatible' content='IE=edge' />" & vbCrLf 
    response.write "<meta http-equiv='cache-control' content='max-age=0' />" & vbCrLf	
    response.write "<meta http-equiv='cache-control' content='no-cache' />" & vbCrLf	
    response.write "<meta http-equiv='expires' content='0' />" & vbCrLf	
    response.write "<meta http-equiv='expires' content='Tue, 01 Jan 1980 1:00:00 GMT' />" & vbCrLf	
    response.write "<meta http-equiv='pragma' content='no-cache' />" & vbCrLf	
		response.write "<link rel='profile' href='http://gmpg.org/xfn/11'>" & vbCrLf				
		response.write "<link rel=""stylesheet"" type=""text/css"" href=""" & CONST_APPLICATION_PATH & "/common/css/css.asp"">" & vbCrLf
    response.write "<script src='common/javascript/jquery.min.js'></script>" & vbCrLf
    response.write "<script src='common/javascript/jquery.numeric.min.js'></script>" & vbCrLf
		response.write "<script src='common/javascript/jquery-ui.min.js'></script>" & vbCrLf
		response.write "<script src='common/javascript/query.switchButton.js'></script>" & vbCrLf
		response.write "<script src='common/javascript/jquery.colorbox-min.js'></script>" & vbCrLf
		response.write "<script src='common/javascript/evacation.js?9'></script>" & vbCrLf	
		response.write "</head>" & vbCrLf
		response.write "<body>" & vbCrLf
		response.write "" & vbCrLf
		response.write "<div id='menuOverlay'>" & vbCrLf
		response.write "   <div id='menutop'>" & vbCrLf
    
		response.write "<table id='menutoptable' style='margin-top:10px'>" & vbCrLf
		response.write "  <tr>" & vbCrLf
		response.write "    <td id='menutopcol1'>" & vbCrLf
		response.write "      <a href='http://shannon.intel.com'><img src='http://shannon.intel.com/wp-content/themes/intel/images/intel-header-blue.png' style='height:50px;vertical-align:middle'></a>" & vbCrLf
		response.write "    </td>" & vbCrLf
		response.write "    <td id='menutopcol2'>" & vbCrLf
		response.write "      <a href='http://shannon.intel.com' class='white' style='margin-right:5px'>Shannon</a> |" & vbCrLf
		response.write "      <a href='http://www.intel.ie' class='white' style='margin-right:5px;margin-left:5px'>Ireland</a> |" & vbCrLf
		response.write "      <a href='https://employeeportal.intel.com' class='white' style='margin-right:5px;margin-left:5px'>Employee Portal</a>		" & vbCrLf
		response.write "    </td>" & vbCrLf
		response.write "    <td id='menutopcol3'>" & vbCrLf
		response.write "" & vbCrLf
		response.write "      <a href='http://shannon.intel.com/contact' class='white'>Feedback</a>" & vbCrLf
		response.write "      <a href='http://shannon.intel.com/contact'>" & vbCrLf
		response.write "         <img style='height:50px;vertical-align:middle;margin-right:5px' src='http://shannon.intel.com/wp-content/themes/intel/images/feedback.png'>" & vbCrLf
		response.write "      </a> " & vbCrLf
		response.write "" & vbCrLf
		response.write "      <a href='http://shannon.intel.com/evacation/default.asp' style='margin-right:10px' class='white'>E-Vacation</a>" & vbCrLf
		response.write "      <a href='http://shannon.intel.com/evacation/default.asp' target='_blank'>" & vbCrLf
		response.write "         <img style='height:50px;vertical-align:middle;margin-right:15px' src='http://shannon.intel.com/wp-content/themes/intel/images/calendar-icon.png'></a> " & vbCrLf
		response.write "" & vbCrLf
		response.write "      <a class='white' href='https://ease.intel.com/es/Phonebook/EditEmployeeRec.aspx' style='margin-right:15px'>" & vbCrLf

    Dim cuser
		Set cuser = new cObjUser
		cuser.SetToLoggedOnUser		

		response.write cuser.FullNameReversed
    
		response.write "      </a>" & vbCrLf
		response.write "      <a class='white' href='https://ease.intel.com/es/Phonebook/EditEmployeeRec.aspx'>" & vbCrLf
		response.write "         <img style='height:50px;vertical-align:middle' src='common/phpldap/userimage.php'></a>" & vbCrLf
		response.write "" & vbCrLf
		response.write "    </td>" & vbCrLf
		response.write "  </tr>			" & vbCrLf
		response.write "</table>" & vbCrLf
    
		response.write "   </div>" & vbCrLf
		response.write "</div>" & vbCrLf
		
		response.write "<div id='welcomeDiv'>" & vbCrLf
		response.write "<div id='welcomeTextContainer'>" & vbCrLf
		response.write "		<h2 class='title'>E-Vacation</h2>" & vbCrLf
		response.write "		<h4 class='description'>Managing your time off...</h4>" & vbCrLf
		response.write "	</div>" & vbCrLf
		response.write "</div>" & vbCrLf

		response.write ""
		response.write "<div id='container'>" & vbCrLf
		response.write "	<div id='content'>" & vbCrLf
		response.write "	"
		response.write "		<div id='wrapper'>" & vbCrLf
		response.write "			<div id='main'>	" & vbCrLf
		
	end function
	

	'** WRITE LEAVE REQUEST FORM ***
	function mWriteLeaveRequestForm(ByRef locobjUser, ByRef locobjLeaveRequest, ByVal locblnValidateRequest)
		Dim objColLeaveTypes
		Dim objCompTime
		Dim strDisabled
		Dim blnDisabled
		Dim loclngCounter
		
		Set objColLeaveTypes = new cColLeaveTypes
		Set objColLeaveTypes.EE = locobjUser
		Set objCompTime = Nothing
		blnDisabled = False
		
		If locobjLeaveRequest.CompTimeID <> 0 Then
		    Set objCompTime = new cObjCompTime
		    objCompTime.ID = locobjLeaveRequest.CompTimeID
		    strDisabled = " readonly"
		    blnDisabled = True
		End If
		
		objColLeaveTypes.CollectionType = CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_EE
		
		If locblnValidateRequest then
			response.write "<tr>"
				response.write "<td style='text-align:center' colspan=4>"
					If not locobjLeaveRequest.FormIsValid then
						mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeaveRequest.FormErrorMessage
					ElseIf not locobjLeaveRequest.NewRequestIsValid then
						mWriteFormError CONST_FORM_ERROR_INVALID, locobjLeaveRequest.NewRequestErrorMessage
					End If
				response.write "</td>"
			response.write "</tr>"
		End If
	
		
		response.write "<br><table class='pageContentHeader'><tr><td><h3>"
			response.write "Leave Request"
		response.write "</h3></td></tr></table>"
		
		response.write "<form name=frmRequestLeave action=""" & CONST_APPLICATION_PATH & "/requestleave.asp"" method=post>"
		
		'**** START HIDDEN FIELD VALUES ****
		response.write "<input type=hidden name=formname value=frmRequestLeave>"
		response.write "<input type=hidden name=ee value="""
			response.write objEEtoView.WWID
		response.write """>"
		
		If Not (objCompTime Is Nothing) Then
			response.Write "<input type=hidden name=m value=" & strMode & ">"
			response.Write "<input type=hidden name=itemid value=" & lngItemID & ">"
			response.Write "<input type=hidden name=fldstrStartTime value=""AM"">" 
			response.Write "<input type=hidden name=fldstrEndTime value=""PM"">" 
		End If
		
		'**** END HIDDEN FIELD VALUES ****	
		
		response.write "<table class='pageContentTable'>"
			response.write "<tr>"
			response.write "<td style='vertical-align:middle'>"
			response.write "<div class='th' style='margin-bottom:10px'>Leave Type</div>"
			response.write "<select class='basicselect' id='leavetype' name='fldstrLeaveType' " & strDisabled & ">"
			
			If Not (objCompTime Is Nothing) then
				response.write "<option value='" & CONST_LEAVE_TYPE_NAME_COMP_TIME & "'>Compensatory Leave</option>"
			Else
				loclngCounter = 0
				while loclngCounter < objColLeaveTypes.Count
					loclngCounter = loclngCounter + 1							    
					if not objColLeaveTypes.Item(loclngCounter).Name = CONST_LEAVE_TYPE_NAME_COMP_TIME then
						response.write "<option value='"
							response.write objColLeaveTypes.Item(loclngCounter).Name & "'"
						if objColLeaveTypes.Item(loclngCounter).Name = locobjLeaveRequest.LeaveType.Name then
							response.write " selected='selected'"
						end if
						response.write ">" & objColLeaveTypes.Item(loclngCounter).Name							        		  
				
						response.write "</option>"
					end if
				wend
			end if
			response.write "</select>"						
			response.write "</td>"						
			response.write "<td style='text-align:right'>"					
			response.write "<div class='th'>Start Date</div>"
			response.write "</td>"
			response.write "<td style='text-align:right;width:200px'>"							
				response.write "<input placeholder='1 Jan 2001' class='evdatepicker' name=fldstrStartDate type=text size=11 maxlength=11 value="""' 
					response.write locobjLeaveRequest.StartDate
				response.write """>"
			response.write "</td>"
			response.write "<td style='text-align:right;width:150px'>"
				mWriteTimeOption "fldstrStartTime", locobjLeaveRequest.StartTime, False							
			response.write "</td>"
			response.write "</tr>"
		
		
			response.write "<tr>"
			response.write "<td style='text-align:left'>"
			
			Dim eloclngCounter
			Dim eloclngResults
			Dim EmployeesAll
			Set EmployeesAll = new cColEmployeesAll				
			
			response.write "<div class='th' style='margin-bottom:10px'>Alternative Approver</div>"
			response.write "  <div class='ui-widget' style='width:100%'>"
			response.write "  <select class='wwidcombobox' name=fldstrApproverWWID>"
			response.write "	<option value=''>Select one...</option>"
			eloclngCounter = 0
			eloclngResults = EmployeesAll.Count
			While eloclngCounter < eloclngResults
				eloclngCounter = eloclngCounter + 1
				with EmployeesAll.Item(eloclngCounter)							
					response.write "<option value='" & .WWID & "'>" & .LastNm & ", " & .FirstNm & " (" & .WWID & ")</option>"
				end with
			Wend
			response.write "</select></div>"					

			response.write "</td>"
			
			response.write "<td style='text-align:right'>"
			response.write "<div class='th'>End Date</div>"
			response.write "</td>"
			
			response.write "<td style='text-align:right;width:200px'>"
				response.write "<input placeholder='1 Jan 2001' class='evdatepicker' name=fldstrEndDate size=11 maxlength=11 type='text' value="""
					response.write locobjLeaveRequest.EndDate
				response.write """>"
			response.write "</td>"
			response.write "<td style='text-align:right;width:150px'>"
				mWriteTimeOption "fldstrEndTime", locobjLeaveRequest.EndTime, blnDisabled 
			response.write "</td>"
		response.write "</tr>"
		
		response.write "<tr>"
			response.write "<td colspan=4>"
				response.write "<br><div class='th' style='margin-bottom:10px'>Comments</div>"
				response.write "<textarea name=fldstrComments maxlength='200' placeholder=""Standard annual leave.."">"'[MFILLAST 08-2006] update to count letters
					response.write mHTMLEncode(locobjLeaveRequest.RequestComments)
				response.write "</textarea><br>"
			response.write "</td>"
		response.write "</tr>"


        Dim locCmd
        Dim locRS        

        Set locCmd = Server.CreateObject("ADODB.Command")
        Set locCmd.ActiveConnection = glbConnection
        locCmd.CommandType = adCmdStoredProc

        locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, trim(objEEtoView.WWID))
        locCmd.CommandText = "usp_get_share_with_team_calendar"
        Set locRS = locCmd.Execute

        if locRs("shareWithTeamCalendar") then
            response.write "<tr>"
                response.write "<td colspan=4><br>"
			        response.write "Add this leave period to team calendar.<br><br>"
			        response.write "<div class='switch-wrapper' style='display:in-line'>"
			        response.write "<input type='checkbox' checked name=fldblnShare value=True>"
			        response.write "</div>"						
			        response.write "</td>"
		    response.write "</tr>"
        end if
        locRS.close
		
		response.write "<tr>"
			response.write "<td colspan=4><br>"
			response.write "Send me an e-mail confirmation of this leave request.<br><br>"
			response.write "<div class='switch-wrapper' style='display:in-line'>"
			response.write "<input type='checkbox'"
			if locobjLeaveRequest.EmailConfOfRequestReq then
				response.write "checked "
			end if
			response.write " name=fldblnNotify value=True>"
			response.write "</div>"	
            response.write "<br><br>"						
		    response.write "<input id='submitleavebutton' type=submit value=""Submit Leave Request""> "					
			response.write "</td>"
        response.write "</tr>"

		response.write "</table>"		
		response.write "</form>"
		response.write "<br><br>"
		
		Set objColLeaveTypes = nothing
		
	end function
	
	
	'*** WRITE COMP TIME INSTANCE DETAILS ****
	function mWriteCompTimeInstanceDetails(ByVal loclngCompTimeID)
	    Dim objCompTime
	    
	    Set objCompTime = new cObjCompTime
	    objCompTime.ID = loclngCompTimeID
	    
		response.write "<table class='pageContentHeader'>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<h3>Compensatory Leave Details</h3>"
					response.write "</td>"
				response.write "</tr>"
		response.write "</table>"	
		
		response.write "<table class='pageContentTable'>"
				response.Write "<tr>"
				    response.Write "<td style='width:175px'>"
				        response.Write "<span class='th'>Date Granted</span>"
				    response.Write "</td>"
				    response.Write "<td style='width:175px'>"
				        response.Write mFormatDate(objCompTime.DateGranted, "medium with day")
				    response.Write "</td>"
				    response.Write "<td style='width:175px'>"
				        response.Write "<span class='th'>Days Available</span>"
				    response.Write "</td>"
				    response.Write "<td style='width:175px'>"
				        response.Write objCompTime.DaysAvailable
				    response.Write "</td>"
				response.Write "</tr>"
				response.Write "<tr>"
				    response.Write "<td>"
				        response.Write "<span class='th'>Expiry Date</span>"
				    response.Write "</td>"
				    response.Write "<td colspan=3>"
				        response.Write mFormatDate(objCompTime.ExpiryDate, "medium with day")
				    response.Write "</td>"
				response.Write "</tr>"
				
				response.Write "<tr>"
				    response.Write "<td>"
				        response.Write "<span class='th'>Reason</span>"
				    response.Write "</td>"
				    response.Write "<td colspan=2>"
				        response.Write objCompTime.Reason
				    response.Write "</td>"
				response.Write "</tr>"
				
		response.write "</table>"
		response.write "<br><br>"
		
		    response.write "<table class='pageContentTable'>"
				response.write "<tr>"
					response.write "<td>"
						    response.write "<br>"
						    response.write "<b>There are no days available for this Comp Time. Please return to the <a href=""" & CONST_APPLICATION_PATH & "/default.asp"">homepage</a>.</b><br>"
						    response.write "<br>"
					response.write "</td>"
				response.write "</tr>"
		    response.write "</table>"
		    response.write "<br><br>"
'		End If
	end function
	
	'*** ERASE COMP TIME INSTANCE DETAILS ****
	function mEraseCompTimeInstanceDetails(ByVal loclngCompTimeID)
	    Dim objCompTime
	    
	    Set objCompTime = new cObjCompTime
	    objCompTime.ID = loclngCompTimeID
	    
		response.write "<table class='pageContentHeader'>"
				response.write "<tr>"
					response.write "<td><h4>"
						response.write "Compensatory Leave Details"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"
			
		response.write "<table class='pageContentTable'>"			
				response.Write "<tr>"
				    response.Write "<td style='width:175px'>"
				        response.Write "<b>Date Granted:</b> "
				    response.Write "</td>"
				    response.Write "<td style='width:175px'>"
				        response.Write mFormatDate(objCompTime.DateGranted, "medium with day")
				    response.Write "</td>"
				    response.Write "<td style='width:175px'>"
				        response.Write "<b>Days Available:</B. "
				    response.Write "</td>"
				    response.Write "<td style='width:175px'>"
				        response.Write objCompTime.DaysAvailable
				    response.Write "</td>"
				response.Write "</tr>"
				response.Write "<tr>"
				    response.Write "<td>"
				        response.Write "<b>Expiry Date:</b> "
				    response.Write "</td>"
				    response.Write "<td colspan=3>"
				        response.Write mFormatDate(objCompTime.ExpiryDate, "medium with day")
				    response.Write "</td>"
				response.Write "</tr>"
				
				response.Write "<tr>"
				    response.Write "<td>"
				        response.Write "<b>Reason:</b> "
				    response.Write "</td>"
				    response.Write "<td colspan=3>"
				        response.Write objCompTime.Reason
				    response.Write "</td>"
				response.Write "</tr>"
				
				response.Write "<tr><td><br></td></tr>"
				
		response.write "</table>"
		response.write "<br><br>"
		
		' Display a error message if there are no available days left in the Comp Time
'		If objCompTime.DaysAvailable = 0 Then
		    response.write "<table class='pageContentTable'>"
				response.write "<tr>"
					response.write "<td style='text-align:center'>"
					    response.write "<center>"
						    response.write "<br>"
						    response.write "<b>There are no days available for this Comp Time. Please return to the <a href=""" & CONST_APPLICATION_PATH & "/default.asp"">homepage</a>.</b><br>"
						    response.write "<br>"
					    response.write "</center>"
					response.write "</td>"
				response.write "</tr>"
		    response.write "</table>"
		    response.write ""
'		End If
	end function
	
	'*** WRITE GRANT LEAVE FORM *** 
	function mWriteGrantLeaveForm(ByRef locobjUser, ByVal locblnMaxCompTimeGranted, ByVal locblnValidateRequest)
		Dim objColLeaveTypes
		Dim locstrDisabled
		Dim loclngCounter		
		
		Set objColLeaveTypes = new cColLeaveTypes
		Set objColLeaveTypes.EE = locobjUser
		
		objColLeaveTypes.CollectionType = CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_EE
		
		If locblnValidateRequest then
					response.write "<table class='pageContentHeader'>"
					response.write "<tr>"
						response.write "<td>"
							If not locobjLeaveRequest.FormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeaveRequest.FormErrorMessage
							ElseIf not locobjLeaveRequest.NewRequestIsValid then
								mWriteFormError CONST_FORM_ERROR_INVALID, locobjLeaveRequest.NewRequestErrorMessage
							End If
						response.write "</td>"
					response.write "</tr>"
					response.write "</table>"
				End If
	
		response.write "<table class='pageContentHeader'"
				response.write "<tr>"
					response.write "<td  class=txttitle  colspan=3>"
						response.write "<h4>Grant Compensatory Leave Form</h4>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"
				
		response.write "<form name=frmGrantCompTime action=""" & CONST_APPLICATION_PATH & "/teamleavesummary.asp?m=gc"" method=post>"
			
				'**** START HIDDEN FIELD VALUES ****
				response.write "<input type=hidden name=formname value=frmGrantCompTime>"
				response.write "<input type=hidden name=ee value="""
					response.write objEEtoView.WWID
				response.write """>"
				'**** END HIDDEN FIELD VALUES ****
				
            response.write "<table class='pageContentTable'>"
				response.write "<tr>"
					response.write "<td><h3>"
						response.write "Employee: "
						response.write locobjUser.FullName
						response.write " (WWID:"
						response.write locobjUser.WWID
						response.write ")"
					response.write "</h3></td>"
				response.write "</tr>"
				
				
				' If Max Comp Days have been granted already print that message
				If locblnMaxCompTimeGranted then				
				    locstrDisabled = " disabled=""disabled"""
				End If 

				If locobjUser.AnnualVacation.CompTime.Count > 0 then
				    response.write "<tr><td>"
				        
				        For loclngCounter = 1 To locobjUser.AnnualVacation.CompTime.Count
				        
				        if locobjUser.AnnualVacation.CompTime.Item(loclngCounter).DaysGranted > 0 then
				            response.write locobjUser.AnnualVacation.CompTime.Item(loclngCounter).DaysGranted & " Day"
				             response.write " Granted (expires "
				            response.write mFormatDate(locobjUser.AnnualVacation.CompTime.Item(loclngCounter).ExpiryDate, "medium with day")
				            response.write ")<br>"
				        end if    
				        
				        Next
				       
				        response.write "</td>"
				        response.write "</tr>"
				End If
				
			    response.write "<tr>"
				    response.write "<td>"
					    response.write "<h4>Days to Grant:&nbsp;"
					    response.write "<select class='basicselect' name=fldlngCompDays" & locstrDisabled & ">"
						    loclngCounter = 0
						  '  while loclngCounter < (CONST_MAX_ANNUAL_COMP_DAYS - locobjUser.AnnualVacation.TotalCompTimeGranted)
							while loclngCounter < CONST_MAX_ANNUAL_COMP_DAYS 
							    loclngCounter = loclngCounter + 1
							    response.write "<option value=""" & loclngCounter & """>" & loclngCounter & "</option>"
						    wend
					    response.write "</select></h4>"
				    response.write "</td>"
			    response.write "</tr>"
			    
			   response.write "<tr>"
				    response.write "<td colspan=3"
				    response.write ">"
					    response.write "<textarea  name=fldstrReason " & locstrDisabled & ">"
					    response.write "</textarea><br>"
				    response.write "</td>"
			    response.write "</tr>"
			    
			    response.write "<tr>"
				    response.write "<td colspan=3 style='text-align:center'>"
					    response.write "<br>"
					    response.write "<input type=submit value=""Grant Leave""" & locstrDisabled & "><br>"
					    response.write "<br>"
				    response.write "</td>"
			    response.write "</tr>"   
			response.write "</table>"
		response.write "</form>"    
		    
		response.write "<br><br>"
		
		Set objColLeaveTypes = nothing
		
	end function
	
	'*** WRITE REVOKE LEAVE FORM *** [CH 12/10/08]
	function mWriteRevokeLeaveForm(ByRef locobjUser, ByVal locblnMinCompTimeGranted, ByVal locblnValidateRequest)
		Dim objColLeaveTypes
		Dim locstrDisabled
		Dim loclngCounter	
		Dim newloclngCounter	
		
		Set objColLeaveTypes = new cColLeaveTypes
		Set objColLeaveTypes.EE = locobjUser
		
		objColLeaveTypes.CollectionType = CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_EE
		
	
		response.write "<table class='pageContentHeader'>"			
				
				response.write "<tr>"
					response.write "<td class=txttitle  colspan=3>"
						response.write "<h4>Revoke Compensatory Leave Form</h4>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"
		
		
		response.write "<form name=frmRevokeCompTime action=""" & CONST_APPLICATION_PATH & "/teamleavesummary.asp?m=gc"" method=post>"	
			
                '**** START HIDDEN FIELD VALUES ****
				response.write "<input type=hidden name=formname value=frmRevokeCompTime>"
				response.write "<input type=hidden name=ee value="""
					response.write objEEtoView.WWID
				response.write """>"
				'**** END HIDDEN FIELD VALUES ****

            response.write "<table class='pageContentHeader'>"
				If locblnValidateRequest then
					response.write "<tr>"
						response.write "<td><h3>"
							If not locobjLeaveRequest.FormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeaveRequest.FormErrorMessage
							ElseIf not locobjLeaveRequest.NewRequestIsValid then
								mWriteFormError CONST_FORM_ERROR_INVALID, locobjLeaveRequest.NewRequestErrorMessage
							End If
						response.write "</h3></td>"
					response.write "</tr>"
				End If

				response.write "<tr>"
					response.write "<td colspan=3>"
						response.write "<h3>Employee: "
						response.write locobjUser.FullName
						response.write " (WWID:"
						response.write locobjUser.WWID
						response.write ")"
					response.write "</h3></td>"
				response.write "</tr>"
				
				' If Min Comp Days have been granted already print that message
				If locblnMinCompTimeGranted then				
				    locstrDisabled = " disabled=""disabled"""
				End If
				
				              
				If locobjUser.AnnualVacation.CompTime.Count > 0 then
				    response.write "<tr>"
				    
				        response.write "<td>"
				        mDebugPrint "comp time: " & locobjUser.AnnualVacation.CompTime.Count		
 				        For loclngCounter = 1 To locobjUser.AnnualVacation.CompTime.Count				           
				            if locobjUser.AnnualVacation.CompTime.Item(loclngCounter).DaysGranted > 0 then
				                response.write locobjUser.AnnualVacation.CompTime.Item(loclngCounter).DaysGranted & " Day"
				                response.write " Granted (expires "
				                response.write mFormatDate(locobjUser.AnnualVacation.CompTime.Item(loclngCounter).ExpiryDate, "medium with day")
				                response.write ")<br>"
				            end if    
				        Next
				        
				        response.write "</td>"
				    response.write "</tr>"
				End If
				
						response.write "<tr>"
				    response.write "<td>"
					    response.write "<h4>Days to Revoke :&nbsp;"
					    
					    response.write "<select class='basicselect' name=fldlngCompDays" & locstrDisabled & ">"
						    loclngCounter = 0
						  
						    while loclngCounter < locobjUser.AnnualVacation.TotalCompTimeGranted
							    loclngCounter = loclngCounter + 1
							
							    response.write "<option value=""" & loclngCounter & """>" & loclngCounter & "</option>"
						    wend	
						   
						    					    
					    response.write "</select></h4>"
				    response.write "</td>"
			    response.write "</tr>"
			    
			    response.write "<tr>"
				    response.write "<td colspan=3>"				    					
						response.write "<textarea name=fldstrReason></textarea><br>"			'										   
				    response.write "</td>"
			    response.write "</tr>"
			    
			    
			    response.write "<tr>"
				    response.write "<td colspan=3 style='text-align:center'>"
					    response.write "<br>"
					    loclngCounter = false
					    response.write "<input type=submit value=""Revoke Leave""" & locstrDisabled & "><br>"					
					    response.write "<br>"
				    response.write "</td>"
			    response.write "</tr>"
			response.write "</form>"    
		response.write "</table>"
		    
		    
		response.write ""
		
		Set objColLeaveTypes = nothing
		
	end function
	
	'**** WRITE LEAVE REQUESTS ****
	function mWriteLeaveRequests(ByRef locobjLeaveRequests, ByVal locblnAdminView)
		
		Dim loclngCounter
		Dim loclngLeaveRequestsCount
		Dim locLeaveRequest
        Dim locCmd
        Dim locRS

        Set locCmd = Server.CreateObject("ADODB.Command")
        Set locCmd.ActiveConnection = glbConnection
        locCmd.CommandType = adCmdStoredProc

        locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, objCurrentUser.WWID)
        locCmd.CommandText = "usp_get_share_with_team_calendar"
        Set locRS = locCmd.Execute     

		response.write "<table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td>"
				response.write "<h3>"
					If locblnAdminView then
						response.write "Leave Periods"
					else
						response.write "Leave Requests"
					end if
				response.write "</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		
		response.write "<table class='pageContentTable'>"
			If locobjLeaveRequests.Count = 0 then
				response.write "<tr>"
					response.write "<td>"
						response.write "No leave requests to display.<br>"
					response.write "</td>"
				response.write "</tr>"
			Else
				response.write "<tr>"
                    response.write "<th>"
						response.write "<br>"
					response.write "</th>"
					response.write "<th>"
						response.write "Leave Type"
					response.write "</th>"
					response.write "<th>"
						response.write "Status"
					response.write "</th>"
					response.write "<th>"
						response.write "Start Date"
					response.write "</th>"
					response.write "<th>"
						response.write "End Date"
					response.write "</th>"
					response.write "<th>"
						response.write "Days"
					response.write "</th>"
					response.write "<th style='text-align:center'>"
                        if locRs("shareWithTeamCalendar") then
                            response.write "Team Calendar"
                        end if
					response.write "</th>"
					response.write "<th>"
						response.write "<br>"
					response.write "</th>"
					response.write "<th>"
						response.write "<br>"
					response.write "</th>"
				response.write "</tr>"
			
			
				loclngCounter = 0
				loclngLeaveRequestsCount = locobjLeaveRequests.Count
				While loclngCounter < loclngLeaveRequestsCount
					loclngCounter = loclngCounter + 1
					With locobjLeaveRequests.Item(loclngCounter)
						response.write "<tr>"
							response.write "<td>"
								response.write "<img src=""" & CONST_APPLICATION_PATH & "/common/images/info.png"" alt='' width=16 height=16 style='border:0px' onclick=""javascript:alert(this.title);"" title=""Request Comments"
									response.write mHTMLEncode(.RequestComments)
									response.write chr(13) & chr(13) & "Response Comments: "
									response.write mHTMLEncode(.ResponseComments)
									if .AwaitingApproval then
										response.write chr(13) & chr(13) & "Awaiting Approval From: "
										response.write .AppointedApprover.FullName
									end if
									response.write chr(13)
					   			response.write """>"
							response.write "</td>"
							response.write "<td>"
								response.write .LeaveType.Name
							response.write "</td>"
							response.write "<td>"
								response.write .Status
							response.write "</td>"
							response.write "<td>"	
								response.write mFormatDate(.StartDate,"medium with day")
								response.write "&nbsp;"
								response.write .StartTime
							response.write "</td>"
							response.write "<td>"
								response.write mFormatDate(.EndDate,"medium with day")
								response.write "&nbsp;"
								response.write .EndTime
							response.write "</td>"
							response.write "<td>"
								response.write .Days
								response.write "&nbsp;&nbsp;&nbsp;"
							response.write "</td>"

                            if locRs("shareWithTeamCalendar")  then
                            response.write "<td style='text-align:left;width:20px;text-align:center'>"							
									if .Status = CONST_LEAVE_PERIOD_STATUS_APPROVED then
                                            If .StartDate > Date then
                                                If .ShareLeaveWithTeamCalendar then
                                                    response.write "<a href=""" & CONST_APPLICATION_PATH & "/teamholidaycalendar.asp?m=htcl&amp;itemid=" & .ID & """ title=""Click here to show/hide this leave request in Team Calendar."">Hide</a>"
                                                else
                                                    response.write "<a href=""" & CONST_APPLICATION_PATH & "/teamholidaycalendar.asp?m=stcl&amp;itemid=" & .ID & """ title=""Click here to show/hide this leave request in Team Calendar."">Show</a>"
                                                End if
                                            end if                                                
									end if
									
									response.write "<br>"
							response.write "</td>"
                            end if

							response.write "<td style='text-align:left'>"
							
						
								if locblnAdminView then
									if .Status = CONST_LEAVE_PERIOD_STATUS_RAISED or _
									   .Status = CONST_LEAVE_PERIOD_STATUS_APPROVED then
										response.write "<a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=lv&amp;ee="
											response.write locobjLeaveRequests.EE.WWID
											response.write "&amp;itemid="
											response.write .ID
										response.write """ title=""View this leave period (with option to delete)."">View</a>"
									end if
								elseif objCurrentUser.WWID = .EE.WWID then								
									if .Status = CONST_LEAVE_PERIOD_STATUS_RAISED or _
									   .Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
									   .Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED then
										response.write "<a href=""" & CONST_APPLICATION_PATH & "/leavesummary.asp?m=cl&amp;itemid="
											response.write .ID
											response.write """ onclick=""return(confirm('Are you sure you want to cancel this leave period (from "
											response.write mFormatDate(.StartDate,"medium with day")
											response.write "&nbsp;"
											response.write .StartTime
											response.write " to "
											response.write mFormatDate(.EndDate,"medium with day")
											response.write "&nbsp;"
											response.write .EndTime
										response.write ")?'))"" title=""Click here to cancel this leave request."">Cancel</a>"
									end if
									
									
								if (.Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
									    .Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED) and _
									   .CanConfirmLeave then
										response.write " <a href=""" & CONST_APPLICATION_PATH & "/leavesummary.asp?m=cf&amp;itemid="
											response.write .ID
											response.write """ onclick=""return(confirm('Click OK to confirm that this leave period was taken: \n\n(from "
											response.write mFormatDate(.StartDate,"medium with day")
											response.write "&nbsp;"
											response.write .StartTime
											response.write " to "
											response.write mFormatDate(.EndDate,"medium with day")
											response.write "&nbsp;"
											response.write .EndTime
										response.write ")'))"" title=""Please click here to confirm that you have taken the leave period"">"
										response.write "Confirm</a>"
									end if
									
								else
									response.write "<br>"
								end if
							response.write "</td>"
							response.write "<td>"
								response.write "<br>"
							response.write "</td>"
						response.write "</tr>"
					End With
				Wend
			end if
			response.write "<tr>"
				response.write "<td colspan=8>"
					response.write "<br>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		response.write ""
	end function
	
	
	'**** WRITE COMP TIME DETAILS ****
	function mWriteCompTimeDetails(ByRef loccolCompTime, ByVal locblnAdminView)
	'**************		
		Dim loclngCounter
		Dim loclngCompTimeCount
		
		response.write "<table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td>"
					response.write "<h3>"
					response.write "Compensatory Time"
					response.write "</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		
		response.write "<table class='pageContentTable'>"
			If loccolCompTime.Count = 0 then
				response.write "<tr>"
					response.write "<td>"
						response.write "No comp time to display.<br><br>"
					response.write "</td>"
				response.write "</tr>"
			Else
				response.write "<tr>"
					response.write "<td>"
						response.write "<br>"
					response.write "</td>"
					response.write "<td style='width:20px'>"
						response.write "<br>"
					response.write "</td>"
					response.write "<td>"
		'				response.write "Status"
					response.write "</td>"
					response.write "<td>"
						response.write "Date Granted"
					response.write "</td>"
					response.write "<td>"
						response.write "Expiry Date"
					response.write "</td>"
					response.write "<td>"
						response.write "Day Granted"
					response.write "</td>"
					response.write "<td>"
						response.write "Day Booked"
					response.write "</td>"
					response.write "<td>"
						response.write "Day Revoked"
					response.write "</td>"
					response.write "<td>"
						response.write "Taken"
					response.write "</td>"
					response.write "<td style='width:50px'>"
						response.write "<br>"
					response.write "</td>"
					response.write "<td>"
						response.write "<br>"
					response.write "</td>"
				response.write "</tr>"
			
				loclngCounter = 0
				loclngCompTimeCount = loccolCompTime.Count
				While loclngCounter < loclngCompTimeCount
					loclngCounter = loclngCounter + 1
					With loccolCompTime.Item(loclngCounter)
						response.write "<tr>"
							response.write "<td>"
								response.write "<br>"
							response.write "</td>"
							response.write "<td>"
								response.write "<img src=""" & CONST_APPLICATION_PATH & "/common/images/info.png"" alt='' width=16 height=16 style='border:0px' onclick=""javascript:alert(this.title);"" title=""Reason: "
									response.write mHTMLEncode(.Reason)
									response.write chr(13) & chr(13) 
								response.write """>"
								
	                    response.write "<td>"
						 response.write .Status
							response.write "</td>"
							response.write "<td>"
								response.write mFormatDate(.DateGranted,"medium with day")
							response.write "</td>"
							response.write "<td>"
								response.write mFormatDate(.ExpiryDate,"medium with day")
							response.write "</td>"
							response.write "<td  align=middle>"
								response.write .DaysGranted
								response.write "&nbsp;&nbsp;&nbsp;"
							response.write "</td>"
							response.write "<td  align=middle>"
								response.write .DaysBooked2
								response.write "&nbsp;&nbsp;&nbsp;"
							response.write "</td>"
							response.write "<td  align=middle>"
								response.write .DaysRevoked2
								response.write "&nbsp;&nbsp;&nbsp;"
							response.write "</td>"
							response.write "<td  align=middle>"
								response.write .Taken
								response.write "&nbsp;&nbsp;&nbsp;"
							response.write "</td>"
							response.write "<td  align=middle>"
                             If .Status = CONST_COMP_TIME_STATUS_EXPIRED  Then
            '                  response.Write "<a href=""" & CONST_APPLICATION_PATH & "/requestleave.asp?m=ct&itemid=" & .ID & """>Book</a>"
                            End If
							response.write "&nbsp;&nbsp;&nbsp;"
							response.write "</td>"
							response.write "<td>"
							
						    If .Status <> CONST_COMP_TIME_STATUS_EXPIRED and .DaysAvailable > 0 Then
                  '            response.Write "<a href=""" & CONST_APPLICATION_PATH & "/requestleave.asp?m=ct&itemid=" & .ID & """>Book</a>"
              '              Response.write  .DaysAvailable   
              '                Response.write  .Days_Booked  
                '              Response.write  .Days_Revoked   
                            End If
							
							response.write "</td>"
							response.write "<td>"
								response.write "<br>"
							response.write "</td>"
						response.write "</tr>"
					End With
				Wend
			end if
		response.write "</table>"
		response.write ""
	end function
	

	'*** WRITE LEAVE SUMMARY ***
	function mWriteLeaveSummary(ByRef locobjUser, ByVal locstrYearSelectorActionURL)
		
        Dim loclngCounter
        Dim loclngCount
        Dim locdatStartDate
        Dim locdatEndDate
        Dim loclngGrantedCompDays
		
        With locobjUser.AnnualVacation
			
        response.write "<table class='pageContentHeader'>"
            response.write "<tr>"
                response.write "<td style='width:auto'>"
                    response.write "<h3>"
                    response.write Trim(UCase(Left(LCase(locobjUser.FirstNm),1)) & Mid(LCase(locobjUser.FirstNm),2))
                    response.write "'s Leave Summary"
                    response.write "</h3>"
                response.write "</td>"	
                response.write "<td style='text-align:right;width:150px'>"	
                    response.write "<span class='th'>Showing Year </span>"
                response.write "</td>"					
                response.write "<td style='text-align:right;width:150px'>"
                    response.write "<form name=""frmYearSelector"" method=post action="""
                    response.write locstrYearSelectorActionURL
                    response.write """>"
                    response.write "<select id='yearselect' style='width:100%' name=lngYear>"
                    mWriteSelectYearOptions(.Year)
                    response.write "</select>"
                    response.write "<input type=hidden name=formname value=""frmYearSelector"">"
                    response.write "<input type=hidden name=m value="""
                    response.write strMode
                    response.write """>"
                    response.write "<input type=hidden name=ee value="""
                    response.write locobjUser.WWID
                    response.write """>"							
                    response.write "</form>"
                response.write "</td>"					
            response.write "</tr>"
        response.write "</table>"
			
        response.write "<table class='pageContentTable' style='width:1000px;'>"
            response.write "<tr>"
                response.write "<td style='width:50px'></td>"
                response.write "<td style='width:280px;vertical-align:top'>"
                    response.write "<img style='height:200px;margin-left:auto;margin-right:auto;margin-top:30px;' src='common/phpldap/userimage.php?user=" & locobjUser.IDSID & "' alt=''>"
                response.write "</td>"
                response.write "<td style='width:450px' valign=top'>"
                    response.write "<div class='th'>Annual Leave</div><br>"
                    response.write "<table>"
                        response.write "<tr>"
                            response.write "<td style='width:400px'>"
                                response.write "Basic Entitlement"
                            response.write "</td>"
                            response.write "<td style='text-align:right;width:50px'>"
							
							If mGetSafeLongInteger(.Year,-1) <> -1 And mGetSafeLongInteger(.Year,-1) <= 2015 then
                                response.write .BasicEntitlementPre2016
							else
								response.write .BasicEntitlement
							End IF
							
                            response.write "</td>"
                        response.write "</tr>"
						
						If mGetSafeLongInteger(.Year,-1) <> -1 And mGetSafeLongInteger(.Year,-1) <= 2015 then
							response.write "<tr>"
								response.write "<td>"
								response.write "Seniority Entitlement"
								response.write "</td>"
								response.write "<td style='text-align:right'>"
								response.write .SeniorityEntitlement
								response.write "</td>"
							response.write "</tr>"
						End if
						
                        loclngGrantedCompDays = 0
                        For loclngCounter = 1 To .CompTime.Count
                        loclngGrantedCompDays = loclngGrantedCompDays + .CompTime.Item(loclngCounter).DaysGranted
                        Next
              
                        response.write "<tr>"
                        response.write "<td>"
                        response.write "Compensatory Leave "
                        response.write "</td>"
                        response.write "<td style='text-align:right'>"
                        response.write loclngGrantedCompDays
                        response.write "</td>"
                        response.write "</tr>"
              
                        response.write "<tr>"
                        response.write "<td>"
                        response.write "Carried Over (from previous year)"
                        response.write "</td>"
                        response.write "<td style='text-align:right'>"
                        response.write .CarryOverEOY
                        response.write "</td>"
                        response.write "</tr>"
              
                        response.write "<tr>"
                        response.write "<td>"
                        response.write "Exception"									
                        response.write "</td>"
                        response.write "<td style='text-align:right'>"
                        response.write .CarryOverPreArranged
                        response.write "</td>"
                        response.write "</tr>"
              
                        If .LegalAdjustmentType <> "" then								
                        response.write "<tr>"
                        response.write "<td>"
                        response.write .LegalAdjustmentType
                        'response.write "&nbsp;Available "
                        response.write "</td>"
                        response.write "<td style='text-align:right'>"
                        response.write .LegalAdjustmentAccrued
                        response.write "</td>"
                        response.write "</tr>"
                        end if             
                        
						
						If mGetSafeLongInteger(.Year,-1) <> -1 And mGetSafeLongInteger(.Year,-1) <= 2015 then						
							response.write "<tr>"
							response.write "<td>"
							response.write "Total Entitlement (Comp Leave NOT Included)"
							response.write "</td>"
							response.write "<td style='border-top: 1px solid #339;text-align:right'>"
							response.write Round(.TotalEntitlementPre2016,2)
							response.write "</td>"
							response.write "</tr>"						
						Else						
							response.write "<tr>"
							response.write "<td>"
							response.write "Total Entitlement (Comp Leave NOT Included)" 
							response.write "</td>"
							response.write "<td style='border-top: 1px solid #339;text-align:right'>"
							response.write Round(.TotalEntitlement,2)
							response.write "</td>"
							response.write "</tr>"						
						End If
              
                        response.write "<tr>"
                        response.write "<td>"
                        response.write "(-) ELP Days Banked"
                        response.write "</td>"
                        response.write "<td style='text-align:right'>"
                        response.write .ELPActive.DaysBankedInCurrentYear + .ELPMatured.DaysBankedInCurrentYear + .ELPUsed.DaysBankedInCurrentYear
                        response.write "</td>"
                        response.write "</tr>"
              
                        response.write "<tr>"
                        response.write "<td>"
                        response.write "(-) Annual Leave Booked"
                        response.write "</td>"
                        response.write "<td style='text-align:right'>"
                        response.write .DaysBooked
                        response.write "</td>"
                        response.write "</tr>"
              
                        response.write "<tr>"
                        response.write "<td>"
                        response.write "Available Balance"
                        response.write "</td>"
                        response.write "<td style='border-top: 1px solid #339;text-align:right'><b>"
						
						If mGetSafeLongInteger(.Year,-1) <> -1 And mGetSafeLongInteger(.Year,-1) <= 2015 then
                           response.write Round(.BalancePre2016,2)
						else
							response.write Round(.Balance,2)
						End IF		
						
                        response.write "</b></td>"
                        response.write "</tr>"	
              
                        response.write "</table>"
					
			            '**** OTHER LEAVE SUMMARY ***
				        With locobjUser.OtherLeave.LeaveGroups
					        response.write "<span class='th' id='showOtherLeave' style='cursor:pointer;'>Other Leave (click to show)</span><br><br>"
					        response.write "<table style='width:450px;display:none' id='otherLeaveTable'>"
						        loclngCounter = 0
						        loclngCount = .Count
						        locdatStartDate = mFirstDayOfYear(locobjUser.YearToView)
						        locdatEndDate = mLastDayOfYear(locobjUser.YearToView)
						        While loclngCounter < loclngCount
							        loclngCounter = loclngCounter + 1
							        response.write "<tr>"
								        response.write "<td style='width:400px'>"
									        response.write .Item(loclngCounter).CollectionLeaveType
								        response.write "</td>"
								        response.write "<td style='width:50px;text-align:right'>"
                                            response.write 0
								        response.write "</td>"
							        response.write "</tr>"
						        Wend
						        If objEEtoView.AnnualVacation.ELPMatured.TargetDays <> "" Then
							        response.Write "<tr>"
							        response.Write "<td></td><td style='border-top: 1px solid #339;text-align:right'>"
								        mWriteMaturedELPBankedDays objEEtoView.AnnualVacation.ELPMatured
							        response.Write "</td>"
							        response.Write "</tr>"
						        End If
					        response.write "</table>"
				        End With

					response.write "</td>"
					response.write "<td style='width:50px'></td>"
					response.write "<td style='width:200px;vertical-align:top;text-align:center'>"
					
					dim daysRemaining
					
					If mGetSafeLongInteger(.Year,-1) <> -1 And mGetSafeLongInteger(.Year,-1) <= 2015 then
                        daysRemaining = Round(.BalancePre2016,2)
					else
						daysRemaining = Round(.Balance,2)
					End If	
					
					If daysRemaining > 14 then
						response.write "<img style='margin-top:60px;' src='common/images/smiley1.png' alt=''>"
					ElseIf daysRemaining > 7 then
						response.write "<img style='margin-top:60px;' src='common/images/smiley2.png' alt=''>"
					else
						response.write "<img style='margin-top:60px;' src='common/images/smiley3.png' alt=''>"
					end if
						
					response.write "<br><br><h3>"
					response.write daysRemaining & " days remaining"
					response.write "</h3>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"
		End With
	end function
	

	'*** WRITE NAV BAR ***	
	function mWriteNavBar(ByVal locstrTabName)
		Dim loclngApprovalsPendingApproval
		Dim loclngLeavePendingConfirmation
       
		loclngLeavePendingConfirmation = objCurrentUser.LeaveRequestsPendingConfirmation

		response.write "<table class='pageContentWidth'>"
			response.write "<tr>"
				response.write "<td>"
				
				With objCurrentUser
				
					response.write "<div class='nav-menu' style='margin-left:0px'><ul>"
			
					'**** HOME PAGE TAB ****					
					response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/default.asp"" title=""Return to the e-Vacation Home Page to see your personal details and user information."">"
						response.write "Home"
					response.write "</a></li>"

									
					'**** TABS FOR EMPLOYEE'S WHOSE LEAVE IS TRACKED ****
					
					if .IsEELeaveTracked then
						response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/userinfo.asp"" title=""See a summary of your user information."">"
					end if
					
					response.write "My Info"
					
					if .IsEELeaveTracked then
						response.write "</a></li>"
					end if
					
					if .IsEELeaveTracked then
						response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/requestleave.asp"" title=""Raise a personal leave request."">"
					end if
					
					response.write "Request Leave"
					
					if .IsEELeaveTracked then
						response.write "</a></li>"
					end if					
					
					
					'**** ELP TAB ****
					if .IsEELeaveTracked then
						response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/elpsummary.asp"" title=""See a summary of your elp entitlement, elp vacation booked, etc."">"
					end if
					
					response.write "ELP"
					
					if .IsEELeaveTracked then
						response.write "</a></li>"
					end if
				

					'**** APPROVAL TAB****
					response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/approverequests.asp"" title=""Approve or reject leave requests and cancellations sent to you for approval."">"
						response.write "Approve Requests"
					response.write "</a></li>"

	
					'**** TABS FOR MANAGERS ****
					if .IsManager then
						response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/teamleavesummary.asp"" title=""View a leave summary for your staff."">"
							response.write "Team Summary"
						response.write "</a></li>"
						
						response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/delegate.asp"" title=""Appoint (or remove) the delegate who approves leave requests raised by your team."">"
							response.write "Delegate"
						response.write "</a></li>"
					end if
	
	
					'**** ADMINISTRATORS TAB ****
					'if .IsAdmin then
					'	response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" title=""Access e-Vacation's administration facilities."">"
					'		response.write "Admin"
					'	response.write "</a></li>"
					'end if
					
					response.write "</ul></div></td>"
					
					response.write "<td style='width:400px;text-align:right'>"
						response.write "<div class='nav-menu'><ul>"						
						response.write "<li><a class='iframe' href='" & CONST_APPLICATION_PATH & "/help.htm' title=""Click here for help with e-Vacation."">"
							response.write "Help"
						response.write "</a></li>"
						response.write "<li><a class='iframe' href='" & CONST_APPLICATION_PATH & "/rulesguide.htm' title=""Click here for e-Vacation's Rules Guide"">"
							response.write "Rules"
						response.write "</a></li>"
						response.write "<li style='margin-right:0px'><a class='iframe' href='" & CONST_APPLICATION_PATH & "/publicholidays.asp' title=""Click here to view Public Holidays recognised by e-Vacation"">"
							response.write "Public Holidays"
						response.write "</a></li>"
						response.write "<li style='margin-right:0px'><a target='_blank' href='" & CONST_APPLICATION_PATH & "/teamholidaycalendar.asp' title=""Click here to view Team Calendar"">"
							response.write "Team Calendar"
						response.write "</a></li>"
					response.write "</ul></div></td>"
				response.write "</tr>"


              response.write "<tr><td><div class='nav-menu' style='margin-left:0px'><ul>"
	
					'**** ADMINISTRATORS TAB ****
					if .IsAdmin then
						response.write "<li><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" title=""Access e-Vacation's administration facilities."">"
							response.write "Admin"
						response.write "</a></li>"
					end if
					
					response.write "</ul></div></td></tr>"
				
				if .IsLeaveApprover and loclngApprovalsPendingApproval > 0 then
					response.write "<tr>"
						response.write "<td align='center' colspan='2'>"
							loclngApprovalsPendingApproval = .ApprovalsPendingApproval
							response.write "<div class='feedbackbox'><b>You currently have "
								response.write loclngApprovalsPendingApproval
							response.write " request"
							if loclngApprovalsPendingApproval <> 1 then
								response.write "s"
							end if
							response.write " pending your approval.</b></div>"
						response.write "</td>"
					response.write "</tr>"
				end if
				
				' Added a message informing users to confirm their leave requests [MOF 10/12/08] 
				if loclngLeavePendingConfirmation > 0 then
					response.write "<tr><td style='text-align:center' colspan='2'>"
						response.write "<b>You have " & loclngLeavePendingConfirmation
						response.write " leave request"
						if loclngLeavePendingConfirmation <> 1 then
							response.write "s"
						end if
						response.write " to cancel/confirm as taken.</b>"
					response.write "</td></tr>"
				end if 
			response.write "</table>"
			response.write ""
		end with
	end function


	'*** WRITE PAGE FOOTER ***
	function mWritePageFooter()		
		response.write "</div>" & vbCrLf
		response.write "		</div>" & vbCrLf
		response.write "	</div>" & vbCrLf
		response.write "</div>" & vbCrLf
		response.write ""
    response.write "<script>" & vbCrLf
		response.write "  (function(i,s,o,g,r,a,m){i['GoogleAnalyticsObject']=r;i[r]=i[r]||function(){" & vbCrLf
		response.write "  (i[r].q=i[r].q||[]).push(arguments)},i[r].l=1*new Date();a=s.createElement(o)," & vbCrLf
		response.write "  m=s.getElementsByTagName(o)[0];a.async=1;a.src=g;m.parentNode.insertBefore(a,m)" & vbCrLf
		response.write "  })(window,document,'script','//www.google-analytics.com/analytics.js','ga');" & vbCrLf
		response.write "" & vbCrLf
		response.write "  ga('create', 'UA-69476019-1', 'auto');" & vbCrLf
		response.write "  ga('send', 'pageview', location.pathname);" & vbCrLf
    
    Dim userparts
    userparts = Split( Request.ServerVariables( "AUTH_USER" ), "\" )
    
    response.write "  ga('set', 'userId', '" & userparts(1) & "');" & vbCrLf
		response.write "" & vbCrLf
		response.write "</script>" & vbCrLf
		response.write "</body></html>" & vbCrLf	
	end function
	
	
	'*** WRITE PUBLIC HOLIDAYS ***
	function mWritePublicHolidays()
		Dim loclngYear
		Dim loclngCounter
		
		loclngYear = 0
		loclngCounter = 0
		
		response.write "<table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td>"
					response.write "<h3>Public Holidays</h3"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"

		response.write "<table class='pageContentTable'>"
			While loclngCounter < glbPublicHolidays.Count
				loclngCounter = loclngCounter + 1
				With glbPublicHolidays.Item(loclngCounter)
					if loclngYear <> datepart("yyyy",.Date) then
						loclngYear = datepart("yyyy",.Date)
						response.write "<tr>"
							response.write "<td>"
								response.write "<br>"
							response.write "</td>"
							response.write "<td style='text-align:right'>"
								response.write "<br><span class='th'>"
								response.write loclngYear
								response.write "</span></b>"
							response.write "</td>"
							response.write "<td colspan=3>"
								response.write "<br>"
							response.write "</td>"
						response.write "</tr>"
					end if
					response.write "<tr>"
						response.write "<td>"
							response.write "<br>"
						response.write "</td>"
						response.write "<td>"
							response.write "<br>"
						response.write "</td>"	
						response.write "<td  align=left>"
							response.write MonthName(month(.Date))
							response.write " "
							response.write Day(.Date)
							response.write " ("
							response.write WeekDayName(weekday(.Date))
							response.write ")"
						response.write "</td>"
						response.write "<td  align=left>"
							response.write mHTMLEncode(.Description)
						response.write "</td>"
						response.write "<td>"
							response.write "<br>"
						response.write "</td>"
					response.write "</tr>"	
				End With
			Wend
			response.write "<tr>"
				response.write "<td  colspan=5>"
					response.write "<br>"
				response.write "</td>"
			response.write "</tr>"	
				
		response.write "</table>"
		response.write ""
		
	end function


	'**** WRITE REPORT FORM ***
	function mWriteReportFormBase()
			
		response.write "<br><table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td>"
					response.write "<h3>Reports</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		
		response.write "<table class='pageContentTable'>"
			response.write "<tr>"
				response.write "<td>"
					response.write "<a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=pr"" title=""View Payroll Report."">"
					Response.Write "Payroll Report"
					Response.Write "</a>"
					response.write "<br><br><a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=hr"" title=""View HR Report."">"
					Response.Write "HR Report"
					Response.Write "</a>"
				response.write "</td>"
			response.write "</tr>"
			
		response.write "</table>"
		response.write "<br><br>"		
	end function
	


	function mWritePayrollReportForm()
		Dim locdatStartDate
		Dim locdatEndDate
			
		response.write "<br><table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td>"
					response.write "<h3>Reports</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		
		response.write "<br><table class='pageContentTable' style='width:450px'>"		
			response.write "<form name=frmReport onSubmit=""javascript:return false;"" action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"
				response.write "<input type=hidden name=frmName value=""frmReport"">"
				response.write "<input type=hidden name=m value="""
					response.write strMode
				response.write """>"
				
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Start Date</span> "
					response.write "</td>"
					response.write "<td>"					
						response.write "<input placeholder='1 Jan 2001' class='evdatepicker' name=fldstrStartDate type=text size=11 maxlength=11 value="""' 
							response.write mFormatDate(locdatStartDate,"medium")
						response.write """>"					
					response.write "</td>"
				response.write "</tr>"

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>End Date</span>"
					response.write "</td>"
					response.write "<td>"
						response.write "<input placeholder='1 Jan 2001' class='evdatepicker' name=fldstrEndDate type=text size=11 maxlength=11 value="""' 
							response.write mFormatDate(locdatEndDate,"medium")
						response.write """>"					
					response.write "</td>"
				response.write "</tr>"

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Employee Status</span>"
					response.write "</td>"
					response.write "<td>"
						response.write "<select class='basicselect' name=""cboStatus"">"
							response.write "<option value=""all"" selected>All</option>"
							response.write "<option value=""active"">Active</option>"
							response.write "<option value=""terminated"">Terminated</option>"
						response.write "</select>"
					response.write "</td>"
				response.write "</tr>"

				response.write "<tr>"

                            response.write "<td>"
						    response.write "<span class='th'>Site</span>"
					        response.write "</td>"
							
							dim cmGetSites
							dim rsSites
							
							set cmGetSites = Server.CreateObject("ADODB.Command")
							set cmGetSites.ActiveConnection = glbConnection
							cmGetSites.CommandType = adCmdStoredProc
							cmGetSites.CommandText = "dbo.pr_evc_shannon_sites"
							set rsSites = cmGetSites.Execute 
							response.write "<td>"
                            Dim fld
                            response.write "<select class='basicselect' name=""cboSite"" style=""width:50px;"">"
							While not rsSites.eof
                                for each fld in rsSites.Fields
                                    response.write "<option value="""
                                    response.write fld
                                    response.write """>"
                                    response.write fld
                                    response.write "</option>"
                                Next
			            
			                    rsSites.movenext
		                    Wend
						    response.write "</select>"
                            response.write "</td>"

							set cmGetSites = nothing
							set rsSites = nothing
							
				response.write "</tr>"
				
				response.write "<tr>"
					response.write "<td></td><td>"
						response.write "<input type=button onclick=""javascript:payrollreport('frmReport','fldstrStartDate','fldstrEndDate','cboStatus','cboSite');"" value=""Generate Report"">"
					response.write "</td>"
				response.write "</tr>"

			response.write "</form>"
			
		response.write "</table>"
		response.write "<br><br>"		
	end function



	function mWriteHRReportForm()
		Dim locdatStartDate
		Dim locdatEndDate
	
		response.write "<br><table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td>"
					response.write "<h3>Reports</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		
		response.write "<br><table class='pageContentTable'>"
									
		response.write "<form name=frmReport onSubmit=""javascript:return false;"" action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"" method=post>"
				response.write "<input type=hidden name=frmName value=""frmReport"">"
				response.write "<input type=hidden name=m value="""
					response.write strMode
				response.write """>"
				
				response.write "<tr>"
					response.write "<td>"
					response.write "<span class='th'>Report Year</span>"
					response.write "</td>"
					response.write "<td>"
					response.write "<input type=hidden  name=""cboMonth""  value=""12"">"
					response.write "<select style='width:100px' class='basicselect' name=""cboYear"">"
				
							Dim intYears
							Dim intCounter
							
							intYears = Year(Date) - 2005
							if intYears <> 0 then
								for intCounter = intYears to 1 step -1
									response.write "<option value=""" 
									Response.Write Year(Date) - intCounter 
									response.write """>"
									Response.Write Year(Date) - intCounter
									response.write "</option>"
								next
							end if
							response.write "<option value=""" 
							Response.Write Year(Date) 
							response.write """ selected>"
							Response.Write Year(Date)
							response.write "</option>"
						Response.Write "</select>"
					response.write "</td>"						
				response.write "</tr>"

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Employee Status</span>"
					response.write "</td>"
					response.write "<td>"
						response.write "<select class='basicselect' name=""cboStatus"">"
							response.write "<option value=""all"" selected>All</option>"
							response.write "<option value=""active"">Active</option>"
							response.write "<option value=""terminated"">Terminated</option>"
						response.write "</select>"
					response.write "</td>"
				response.write "</tr>"
    
                            response.write "<td>"
						    response.write "<span class='th'>Site</span>"
					        response.write "</td>"
							
							dim cmGetSites
							dim rsSites
							
							set cmGetSites = Server.CreateObject("ADODB.Command")
							set cmGetSites.ActiveConnection = glbConnection
							cmGetSites.CommandType = adCmdStoredProc
							cmGetSites.CommandText = "dbo.pr_evc_shannon_sites"
							set rsSites = cmGetSites.Execute 
							response.write "<td>"
                            Dim fld
                            response.write "<select class='basicselect' name=""cboSite"" style=""width:50px;"">"
							While not rsSites.eof
                                for each fld in rsSites.Fields
                                    response.write "<option value="""
                                    response.write fld
                                    response.write """>"
                                    response.write fld
                                    response.write "</option>"
                                Next
			            
			                    rsSites.movenext
		                    Wend
						    response.write "</select>"
                            response.write "</td>"


							set cmGetSites = nothing
							set rsSites = nothing
		
						response.write "<input type=hidden  name=""cboExempt"" value=""all"">"
				
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Employee WWID</span>"
					response.write "</td>"
					response.write "<td>"
					
						Dim eloclngCounter
						Dim eloclngResults
						Dim EmployeesAll
						Set EmployeesAll = new cColEmployeesAll				
						
						response.write "  <div class='ui-widget' style='width:100%'>"
						response.write "  <select class='wwidcombobox' name=txtEmployee>"
						response.write "	<option value=''>Select one...</option>"
						eloclngCounter = 0
						eloclngResults = EmployeesAll.Count
						While eloclngCounter < eloclngResults
							eloclngCounter = eloclngCounter + 1
							with EmployeesAll.Item(eloclngCounter)											
								response.write "<option value='" & .WWID & "' "
								response.write ">" & .LastNm & ", " & .FirstNm & " (" & .WWID & ")</option>"
							end with
						Wend
						response.write "</select></div>"				
						
					response.write "</td>"
				response.write "</tr>"

				response.write "<tr>"
				response.write "<td>"
					response.write "<span class='th'>Manager WWID</span>"
				response.write "</td>"
				response.write "<td  align=left>"
				
					Set EmployeesAll = new cColEmployeesAll				
					
					response.write "  <div class='ui-widget' style='width:100%'>"
					response.write "  <select class='wwidcombobox' name=txtManager>"
					response.write "	<option value=''>Select one...</option>"
					eloclngCounter = 0
					eloclngResults = EmployeesAll.Count
					While eloclngCounter < eloclngResults
						eloclngCounter = eloclngCounter + 1
						with EmployeesAll.Item(eloclngCounter)											
							response.write "<option value='" & .WWID & "' "
							response.write ">" & .LastNm & ", " & .FirstNm & " (" & .WWID & ")</option>"
						end with
					Wend
					response.write "</select></div>"	
				
					response.write "</td>"	
				response.write "</tr>"
																				
				response.write "<tr>"
				response.write "<td></td><td>"
				response.write "<input type=button onclick=""javascript:hrreport('frmReport','cboMonth','cboYear','cboExempt','cboSite','cboStatus','txtEmployee','txtManager');"" value=""Generate Report"" id=button1 name=button1>"
				response.write "</td>"
				response.write "</tr>"

			response.write "</form>"
			
		response.write "</table><br><br>"
		response.write ""		
	end function
	
	
	'*** WRITE SELECT YEAR OPTIONS ***
	function mWriteSelectYearOptions(ByVal loclngYear)
		Dim loclngCounter
		Dim loclngStartYear
		Dim loclngEndYear
		
		loclngYear = mGetSafeLongInteger(loclngYear,-1)
		
		if loclngYear = -1 then loclngYear = Year(date)
		
		loclngStartYear = loclngYear - 5
		loclngEndYear = loclngYear + 5

		if loclngStartYear < CONST_FIRST_YEAR_SYSTEM_ACTIVE then
			loclngStartYear = CONST_FIRST_YEAR_SYSTEM_ACTIVE
		end if
		
		if loclngStartYear > datepart("yyyy",date()) then
			loclngStartYear = datepart("yyyy",date())
		end if
		
		if loclngEndYear < loclngStartYear + 5 then
			loclngEndYear = loclngStartYear + 5
		end if
		
		for loclngCounter = loclngStartYear to loclngEndYear
			response.write "<option value="
			response.write loclngCounter
			if loclngYear = loclngCounter then
				response.write " selected"
			end if
			response.write ">"
			response.write loclngCounter
		next
	end function
	
	
	'*** WRITE TEAM LEAVE SUMMARY ***
	function mWriteTeamLeaveSummary(ByRef locobjUser)
		Dim loclngIndex
		Dim loclngCountReports
		Dim locstrClassText
		
		If Request.Form("update") = "yes" Then
		    updateEmployees objEEtoView
		End if

		response.write "<table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td><h3>"
					response.write "Team Leave Summary"
				response.write "</h3></td>"		
			response.write "</tr>"	
		response.write "</table>"
		
		response.write "<table class='pageContentTable'>"
			With objCurrentUser.DirectReports
				loclngCountReports = .Count
				If loclngCountReports = 0 then
				response.write "<tr>"
					response.write "<td style='text-align:center'>"
						response.write "No employees were found which report directly to you.<br>"
					response.write "</td>"
				response.write "</tr>"
				Else
					response.Write "<tr>"
						response.Write "<td style='text-align:right' colspan=7>"							
							response.write "<div class='nav-menu'><ul>"							
							response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/holidaycalendar.asp"" title=Click here to view the calendar of your employee holidays. class=navLink>Monthly Calendar</a>"
							response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/yearlycalendar.asp"" title=Click here to view a years display of employee holidays. class=navLink>Calendar Year</a>"
							response.write "<li style='margin-left:0px'><a href=""" & CONST_APPLICATION_PATH & "/caldisplay.asp"" title=Click here to choose what time span to view of employee holidays. class=navLink>Choose Time Period</a>"
							response.write "</ul></div>"
						response.Write "</td>"
						response.Write "<td style='text-align:right' colspan=7>"	
%>
<form name="form1" action="teamleavesummary.asp" method="post">
    <input type="hidden" name="update" value="yes">
    <input type="submit" value="Update Employee List" id='employeeListButton'>
</form>

<%								
						response.Write "</td>"
					response.Write "</tr>"
					response.write "<tr>"
						response.write "<td><span class='th'>"
							response.write "Name"
						response.write "</span></td>"
						response.write "<td><span class='th'>"
							response.write "WWID"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Total Entitlement for " & DatePart("yyyy", Date) & """><span class='th'>"
							response.write "Entitlement"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Annual Leave Balance"" style=""border-left: 1px solid #AAAAFF""><span class='th'>"
							response.write "A/L<br>Balance"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Days banked into Extended Leave Program""><span class='th'>"
							response.write "ELP"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Compensatory Days available to Employee""><span class='th'>"
							response.write "Comp<br>Leave"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Leave Days Pending Your Approval""><span class='th'>" ' [MOF 1/09] 
							response.write "Pending<br>Approval"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Leave Days that have been Approved""><span class='th'>" ' [MOF 1/09] 
							response.write "Approved"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Leave Days Booked (Not Expired)"" style=""border-left: 1px solid #AAAAFF;""><span class='th'>" ' [MOF 1/09] 
							response.write "Booked"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Leave Days an Employee should Cancel/Confirm as Taken""><span class='th'>" ' [MOF 1/09] 
							response.write "Expired"
						response.write "</span></td>"
						response.write "<td  style='text-align:right' title=""Leave Days Employee has Confirmed as Taken""><span class='th'>" ' [MOF 1/09] 
							response.write "Taken"
						response.write "</span></td>"
						response.write "<td style='text-align:right;border-left: 1px solid #AAAAFF'>" ' [MOF 10/12/08] 
							response.write "<br>"
						response.write "</td>"
						response.write "<td>"
							response.write "<br>"
						response.write "</td>"
						response.write "<td>"
							response.write "<br>"
						response.write "</td>"
					response.write "</tr>"

								
					loclngIndex = 0
					While loclngIndex < loclngCountReports
						loclngIndex = loclngIndex + 1
						With .Item(loclngIndex)
							If .IsEELeaveTracked then
								locstrClassText = ""
							else
								locstrClassText = " class=txtdim"
							end if
							response.write "<tr>"
								response.write "<td"
									response.write locstrClassText
								response.write ">"
									response.write .FullName
									if .ActiveStatus <> "A" then
										response.write "<br>("
										response.write .ActiveStatusName
										response.write ")"
									end if
								response.write "</td>"
								response.write "<td "
									response.write locstrClassText
								response.write ">"
									response.write .WWID
								response.write "</td>"
								
								response.Write "<td style='text-align:right'>"
								    response.Write Round(.AnnualVacation.TotalEntitlement,2)
								response.Write "</td>"
								
								response.write "<td style='text-align:right;border-left: 1px solid #AAAAFF'"
									response.write locstrClassText
								response.write ">"
									If .IsEELeaveTracked then
										response.write .AnnualVacation.BalanceForMgrView
									else
										response.write "<br>"
									end if
								response.write "</td>"
								
								response.Write "<td style='text-align:right'>"
								    response.Write .AnnualVacation.ELPActive.DaysBankedInCurrentYear + _
								                   .AnnualVacation.ELPMatured.DaysBankedInCurrentYear + _
								                   .AnnualVacation.ELPUsed.DaysBankedInCurrentYear
								response.Write "</td>"
								
								response.Write "<td style='text-align:right'>"
								    response.Write .AnnualVacation.AvailableCompTime
								response.Write "</td>"
																
								response.write "<td style='text-align:right'>"
									response.write .AnnualVacation.DaysPendingApproval ' [MOF 1/09]
								response.write "</td>"
								
								response.write "<td style='text-align:right'>"
									response.write .AnnualVacation.DaysApproved ' [MOF 1/09]
								response.write "</td>"
								
								response.write "<td align=right style=""border-left: 1px solid #AAAAFF"">"
									response.write .AnnualVacation.DaysBookedNotExpired ' [MOF 1/09]
								response.write "</td>"
								
								response.write "<td style='text-align:right'>"
									response.write .DaysToConfirm ' [MOF 1/09]
								response.write "</td>"
								
								response.write "<td style='text-align:right'>"
									response.write .DaysConfirmed ' [MOF 10/12/08]
								response.write "</td>"
								
								response.write "<td align=right" ' [MOF 1/09]
									response.write locstrClassText
								response.write " style=""border-left: 1px solid #AAAAFF"">"								
								
								If .IsEELeaveTracked then
									response.write "<a href=""" & CONST_APPLICATION_PATH & "/grantleave.asp?ee="
									response.write .WWID
									response.write """ title=""Click here to grant compensatory days for this employee."" class=navLink>Grant</a>"														
								else
									response.write "<br>"
								end if
									
								response.write "<td " 
									response.write locstrClassText
								response.write " style='text-align:right;border-left: 0px solid #AAAAFF'>"
								
								If .IsEELeaveTracked then
									response.write "<a href=""" & CONST_APPLICATION_PATH & "/revokeleave.asp?ee="
									response.write .WWID
									response.write """ title=""Click here to revoke compensatory days for this employee."" class=navLink>Revoke</a>"														
								else
									response.write "<br>"
								end if									
									
								response.write "</td>"	
								
								response.write "<td style='text-align:right' "
									response.write locstrClassText
								response.write ">"
									If .IsEELeaveTracked then
										response.write "<a href=""" & CONST_APPLICATION_PATH & "/leavesummary.asp?m=el&amp;ee="
										response.write .WWID
										response.write """ title=""Click here to display more details for this employee."" class=navLink>Details</a>"
									else
										response.write "<br>"
									end if
								response.write "</td>"
							response.write "</tr>"
						End With
					Wend
				End If
			End With
			
		response.write "</table>"
		response.write "<br>"
		
		'***DISPLAY DESCRIPTION OF TABLE
		response.write "<table class='pageContentTable'>"
			response.write "<tr>"
				response.write "<td style='text-align:center'>"
					response.write "<br>"
					response.write "<b>Key:</b>&nbsp;&nbsp;&nbsp;&nbsp;<span class=txtdim>Leave Not Tracked On e-Vacation</span>&nbsp;&nbsp;&nbsp;&nbsp;Leave Tracked On e-Vacation<br>"
					response.write "<br>"
				response.write "</td>"
			response.write "</tr>"
			response.write "<tr>"
				response.write "<td style='text-align:center'>"
					response.write "<b>Entitlement</b> = A/L Balance + ELP + Comp Time + Pending Approval + Approved<br>"
					response.write "<br>"
					response.write "<b>Approved</b> = Booked + Expired + Taken<br>"
					response.write "<br>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
		response.write ""
		
		response.write "<table class='pageContentHeader'>"
			response.write "<tr>"
				response.write "<td><h3>"
					response.write "Direct Reports Reports"
				response.write "</h3></td>"
			response.write "</tr>"	
		response.write "</table>"
		
		'**CHECKS FOR EMPLOYEES EMPLOYEEs
		
		Dim rstEmployeesEmps
		Dim cmGetEmployeesEmployeeData
		Dim m_cnDB
		
		With objCurrentUser.DirectReports
				loclngCountReports = .Count
				If loclngCountReports = 0 then
				
				Else
					loclngIndex = 0
					While loclngIndex < loclngCountReports
						loclngIndex = loclngIndex + 1
						With .Item(loclngIndex)
							If .IsEELeaveTracked then
								locstrClassText = ""
							else
								locstrClassText = " class=txtdim"
							end if
							
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
							Set cmGetEmployeesEmployeeData = Server.CreateObject("ADODB.Command")
							Set cmGetEmployeesEmployeeData.ActiveConnection =  m_cnDB
							cmGetEmployeesEmployeeData.CommandType = 4
							cmGetEmployeesEmployeeData.CommandText = "dbo.team_emps_emps"
							cmGetEmployeesEmployeeData.Parameters.Append cmGetEmployeesEmployeeData.CreateParameter("@vWWID", adChar, adParamInput, 8, .WWID)
							Set rstEmployeesEmps = cmGetEmployeesEmployeeData.Execute
							
							if not rstEmployeesEmps.eof then	
		                        response.write "<table class='pageContentTable'>"
							        response.write "<tr>"
								        response.write "<td style='width:350px'>"
									        response.write .FullName
									        if .ActiveStatus <> "A" then
										        response.write " ("
										        response.write .ActiveStatusName
										        response.write ")"
									        end if
								        response.write "</td>"
								        response.write "<td style='width:200px'>"
									        response.write .WWID
								        response.write "</td>"        								
								        response.write "<td align='left'>"
									        If .IsEELeaveTracked then
										        response.write "<a href=""" & CONST_APPLICATION_PATH & "/leavesummary.asp?m=el&amp;ee="
										        response.write .WWID
										        response.write """ title=""Click here to display more details for this employee."" class=navLink>Details</a>"
									        else
										        response.write "<br>"
									        end if
								        response.write "</td>"
							        response.write "</tr>"								
							        
							    Do while not rstEmployeesEmps.eof
							        response.Write "<tr>"
							            response.Write "<td style='padding-left:25px'>"
    							                response.Write rstEmployeesEmps.fields.item(1).value
		    					                response.write " "
	    						                response.Write rstEmployeesEmps.fields.item(2).value
							            response.Write "</td>"
							            response.write "<td>"
									        response.write rstEmployeesEmps.fields.item(0).value
								        response.write "</td>" 
							            response.Write "<td>"
								            response.write "<a href=""" & CONST_APPLICATION_PATH & "/leavesummary.asp?m=el&amp;ee="
								            response.write rstEmployeesEmps.fields.item(0).value
								            response.write """ title=""Click here to display more details for this employee."" class=navLink>Details</a>"
							            response.Write "</td>"    							
							        response.Write "</tr>"
        							
							        rstEmployeesEmps.movenext
							    loop
        							
						        response.Write "</table>"    
		                        response.write "<br><br>"
							end if  
						End With
					Wend 
				End If
			End With	
		
	end function
	
	
	'*** WRITE TIME OPTION ***
	function mWriteTimeOption(ByVal locstrFieldName, ByVal locstrTime, ByVal locblnDisabled)
		response.write "<select style='width:100%' class='ampmselect' name="
			response.write locstrFieldName
			if locblnDisabled then
				response.write " disabled"
			end if
		response.write ">"
			response.write "<option value=""AM"""
				if locstrTime = "AM" then
					response.write " selected"
				end if
			response.write ">AM"
			response.write "<option value=""PM"""
				if locstrTime = "PM" then
					response.write " selected"
				end if
			response.write ">PM"
		response.write "</select>"
	end function
	

	'*** WRITE USER ERROR ***
	function mWriteUserError(ByVal loclngErrorCode)
		'CA Fixed Error handling to call user error page instead of NewUsers.asp page
		'for CONST_LOGON_ERROR_USER_SET_UP_REQUIRED, 22 May 2001	
		response.write "<br><table class='pageContentHeader'><tbody><tr><td><h3>"
		response.write CONST_USER_ERROR_TITLE(loclngErrorCode)	
		response.write "</h3></td></tr></tbody></table>"
		response.write "<table class='pageContentTable'>"
			response.write "<tr>"
				response.write "<td>"
					response.write "<br><b>"
					Select Case loclngErrorCode
						Case CONST_LOGON_ERROR_USER_BLANK '1
							response.write "Sorry - the system was unable to verify your login ID."
						Case CONST_LOGON_ERROR_USER_NOT_FOUND '2
							response.write "Sorry - your login ID was not found in the Worker Data Services Database 'WDS'.<br><br>" &_
											"There may be a problem with your browser configuration, your network connection,<br>" & _
											"or the information held in 'WDS' may be incorrect or out of date."
						Case CONST_LOGON_ERROR_USER_ACCESS_DENIED '3
							response.write "Sorry - access to e-Vacation has been denied."
							response.write "<br><br>" 
														
						Case CONST_LOGON_ERROR_USER_SET_UP_REQUIRED '4
							Response.Write "Sorry - please contact e-Vacation Administrator to grant you access to e-Vacation.<br>"
							response.write "<br>"
							
						Case CONST_USER_PAGE_ACCESS_DENIED '5
							response.write "Sorry - you do not have access to the page requested.<br>"
							response.write "<br>"
							response.write "Click <a href=""" & CONST_APPLICATION_PATH & "/default.asp"" class=error>here</a> to continue.</a>"
						Case CONST_USER_NOT_ALLOWED_TO_DELEGATE '6
							response.write "Sorry - access to the 'Delegate' feature is disabled as you do not appear to have any direct reports.<br>"
							response.write "<br>"
							response.write "Click <a href=""" & CONST_APPLICATION_PATH & "/default.asp"" class=error>here</a> to continue.</a>"
						Case Else
							response.write "Sorry - an unknown error has occurred while attempting to verify your identity."
							response.write "Click <a href=""" & CONST_APPLICATION_PATH & "/default.asp"" class=error>here</a> to return to e-Vacation.</a>"
					End Select
					response.write "</b><br><br><br><br><br><br>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
	end function
	
	
	'*** WRITE VIEWING EMPLOYEE ***
	function mWriteViewingEmployee(ByRef locobjUser)
		if locobjUser.WWID <> objCurrentUser.WWID then
			response.write "<table class='pageContentHeader'>"
				response.write "<tr>"
					response.write "<td><h3>"
						response.write "Viewing: "
						response.write locobjUser.FullName
						response.write " (WWID:"
						response.write locobjUser.WWID
						response.write ")"
					response.write "</h3></td>"
				response.write "</tr>"
			response.write "</table>"
		end if
	end function
	
	'*** WRITE ELP VACATION DETAILS ***
	function mWriteELPVacationDetails(ByRef objUser, ByRef locobjUsedELP, ByVal isAdmin)
		With locobjUsedELP.LeavePeriod
			
			response.write "<table class='pageContentHeader'>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<h3>ELP Vacation</h3>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"
			
			response.write "<table class='pageContentTable'>"
				response.write "<tr>"
					response.write "<td style='width:250px'>"
						response.write "<span class='th'>Start Date</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(.StartDate,"medium with day") & " (" & .StartTime & ")"
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>End Date</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(.EndDate,"medium with day") & " (" & .EndTime & ")"
					response.write "</td>"
				response.write "</tr>"
					
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Total Days</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .Days
					response.write "</td>"
				response.write "</tr>"

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Status</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .Status
					response.write "</td>"
				response.write "</tr>"
				
				response.write "<tr>"
					response.write "<td></td>"
					response.write "<td>"
						if .Status = CONST_LEAVE_PERIOD_STATUS_RAISED or .Status = CONST_LEAVE_PERIOD_STATUS_APPROVED then
							if isAdmin then
								response.write "<a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=elp&ee="
								response.write objUser.WWID
								response.write "&strMode=cl&amp;itemid="
							else
								response.write "<a href=""" & CONST_APPLICATION_PATH & "/elpsummary.asp?m=cl&amp;itemid="
							end if
							response.write .ID
							response.write """ onclick=""return(confirm('Are you sure you want to cancel this ELP leave period (from "
							response.write mFormatDate(.StartDate,"medium with day")
							response.write "&nbsp;"
							response.write .StartTime
							response.write " to "
							response.write mFormatDate(.EndDate,"medium with day")
							response.write "&nbsp;"
							response.write .EndTime
							response.write ")?'))"" title=""Click here to cancel this ELP leave request."">Cancel</a>"
						end if
					response.write "</td>"
				response.write "</tr>"									
			response.write "</table>"
			response.write "<br><br>"
		end with
	end function
	
	'*** WRITE MATURED ELP BANKED DAYS
	function mWriteMaturedELPBankedDays(ByRef locobjELP)
		Dim loclngYear
		Dim loclngEndYear
		With locobjELP
			response.Write "<tr>"
				response.write "<td>"
					response.write "<b>ELP Days available to take:</b>"
				response.write "</td>"
				response.write "<td colspan='100%' align='right'>"
					response.write .TargetDays
				response.write "</td>"
			response.Write "</tr>"
		end with
	end function
	
	
	'*** WRITE MATURED ELP DETAILS ***
	function mWriteMaturedELPDetails(ByRef locobjELP, ByVal AdminView)
		Dim loclngYear
		Dim loclngEndYear
		With locobjELP
		
			response.write "<br><br><table class='pageContentHeader'>"
				response.write "<tr>"
					response.write "<td><h3>"
						response.write .Status
						response.write " ELP"
					response.write "</h3></td>"
				response.write "</tr>"	
			response.write "</table>"
			
			response.write "<table class='pageContentTable'>"
			
				response.write "<tr>"
					response.write "<td style='width:250px'>"
						response.write "<span class='th'>Activated On</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(.DateActivated,"medium with day")
					response.write "</td>"
				response.write "</tr>"
				
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Activated by</span>"
					response.write "</td>"
					response.write "<td>"
						response.write .ActivatedBy.FullName
						response.write " ("
						response.write .ActivatedBy.WWID
						response.write ")"
					response.write "</td>"
				response.write "</tr>"
					
				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Days available to take</span>"
					response.write "</td>"
					response.write "<td  colspan=2>"
						response.write .TargetDays
					response.write "</td>"
				response.write "</tr>"

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Expire"
							If .ExpiryDate <= date() then
								response.write "d"
							Else
								response.write "s"
							End If
						response.write "</span>"
					response.write "</td>"
					response.write "<td>"
						response.write mFormatDate(.ExpiryDate,"medium")
					response.write "</td>"
				response.write "</tr>"
				
				If AdminView then
					response.write "<tr>"
						response.write "<form method=get action=""" & CONST_APPLICATION_PATH & "/adminhome.asp"">"
							response.write "<td colspan='100%'><br>"
								response.write "<input type=hidden name=m value=""eelp"">"
								response.write "<input type=hidden name=ee value="""
									response.write objEEtoView.WWID
								response.write """>"
								response.write "<input type=hidden name=itemid value="""
									response.write .ELPID
								response.write """>"
								response.write "<input type=submit value=""Edit ELP Activation Details"" title=""Edit the activation details for this ELP instance."">"
						response.write "</td>"
						response.write "</form>"
					response.write "</tr>"
				End If
									
			response.write "</table>"
			response.write "<br><br>"
		end with
	end function
	
	
	function mWriteELPLeaveRequestForm(ByRef objEEtoView, ByRef locobjLeaveRequest, ByVal locblnValidateRequest)
		Dim loclngCounter
			
		response.write "<table class='pageContentHeader'>"		
			response.write "<tr>"
				response.write "<td>"
					response.write "<h3>ELP Leave Request Form</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"	
		
		response.write "<table class='pageContentTable'>"		
			response.write "<form name=frmRequestELPLeave action=""" & CONST_APPLICATION_PATH & "/elpsummary.asp"" method=post "
			response.write "onSubmit=""Javascript:alert('Please consider booking annual leaves in addition to your Extended Leave to facilitate your replacement.');"""
			response.write ">"

				'**** START HIDDEN FIELD VALUES ****
				response.write "<input type=hidden name=formname value=frmRequestELPLeave>"
				response.write "<input type=hidden name=ee value="""
					response.write objEEtoView.WWID
				response.write """>"
				response.write "<input type=hidden name=fldlngELPID value="""
					response.write objEEtoView.AnnualVacation.ELPMatured.ELPID
				response.write """>"
				'**** END HIDDEN FIELD VALUES ****			

				
				If locblnValidateRequest then
					response.write "<tr>"
						response.write "<td style='text-align:center' colspan=4>"
							If not locobjLeaveRequest.FormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeaveRequest.FormErrorMessage
							ElseIf not locobjLeaveRequest.NewRequestIsValid then
								mWriteFormError CONST_FORM_ERROR_INVALID, locobjLeaveRequest.NewRequestErrorMessage
							End If
						response.write "</td>"
					response.write "</tr>"
				End If

				response.write "<tr>"
					response.write "<td>"
						response.write "<span class='th'>Start Date</span>" 
					response.write "</td>"
				
					response.write "<td  width=30% align=left>"
						
						response.write "<input placeholder='1 Jan 2001' class='evdatepicker' name=fldstrStartDate type=text size=11 maxlength=11 value="""' 
							response.write locobjLeaveRequest.StartDate
						response.write """>"
						
					response.write "</td>"
					
					response.write "<td>"
						response.write "<span class='th'>Alternative Approver</span>" 
					response.write "</td>"
				
					response.write "<td align=left "
						if locblnValidateRequest then
							if locobjLeaveRequest.IsInvalidField("Approver.WWID") then
								response.write " class=error"
							end if
						end if
					response.write ">"
					
						Dim eloclngCounter
							Dim eloclngResults
							Dim EmployeesAll
							Set EmployeesAll = new cColEmployeesAll				
							
							response.write "  <div class='ui-widget' style='width:100%'>"
							response.write "  <select class='wwidcombobox' name=fldstrApproverWWID>"
							response.write "	<option value=''>Select one...</option>"
							eloclngCounter = 0
							eloclngResults = EmployeesAll.Count
							While eloclngCounter < eloclngResults
								eloclngCounter = eloclngCounter + 1
								with EmployeesAll.Item(eloclngCounter)											
									response.write "<option value='" & .WWID & "' "
									response.write ">" & .LastNm & ", " & .FirstNm & " (" & .WWID & ")</option>"
								end with
							Wend
							response.write "</select></div>"	
						
					response.write "</td>"
				
				response.write "</tr>"
			
				response.write "<tr>"
					response.write "<td  colspan=4"
						if locblnValidateRequest then
							if locobjLeaveRequest.IsInvalidField("RequestComments") then
								response.write " class=error"
							end if
						end if
					response.write ">"
						response.write "<div class='th' style='margin-bottom:10px'>Comments</div>"
						response.write "<textarea name=fldstrComments maxlength='200'>"
							response.write mHTMLEncode(locobjLeaveRequest.RequestComments)
						response.write "</textarea><br>"
						
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td colspan=4>"
						response.write "Send me an e-mail confirmation of this ELP leave request<br><br>"
						response.write "<input type=checkbox "
							if locobjLeaveRequest.EmailConfOfRequestReq then
								response.write "checked "
							end if
						response.write "name=fldblnNotify value=True>"
						response.write ""
					response.write "</td>"
				response.write "</tr>"
				response.write "<tr>"
					response.write "<td colspan=4>"
						response.write "<br>"
						response.write "<input type=submit value=""Submit ELP Leave Request""><br>"
						response.write "<br>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</form>"
		response.write "</table>"
		response.write "<br><br>"
		
			
	end function

	function mWriteELPAdminLeaveRequestForm(ByRef objEEtoView, ByRef locobjLeaveRequest, ByVal locblnValidateRequest)
		Dim loclngCounter
			
		response.write "<table class='pageContentHeader'>"		
			response.write "<tr>"
				response.write "<td>"
					response.write "<h3>ELP Leave Request Form</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"	
		
		response.write "<form name=frmRequestELPLeave action=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=elp&ee="
		response.write objEEtoView.WWID
		response.write """ method=post "
		response.write "onSubmit=""Javascript:alert('Please consider booking annual leaves in addition to your Extended Leave to facilitate your replacement.');"""
		response.write ">"

			'**** START HIDDEN FIELD VALUES ****
			response.write "<input type=hidden name=formname value=frmRequestELPLeave>"
			response.write "<input type=hidden name=ee value="""
				response.write objEEtoView.WWID
			response.write """>"
			response.write "<input type=hidden name=fldlngELPID value="""
				response.write objEEtoView.AnnualVacation.ELPMatured.ELPID
			response.write """>"
			'**** END HIDDEN FIELD VALUES ****
			
			response.write "<table class='pageContentTable'>"
				If locblnValidateRequest then
					response.write "<tr>"
						response.write "<td colspan=2>"
							If not locobjLeaveRequest.FormIsValid then
								mWriteFormError CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY, locobjLeaveRequest.FormErrorMessage
							ElseIf not locobjLeaveRequest.NewRequestIsValid then
								mWriteFormError CONST_FORM_ERROR_INVALID, locobjLeaveRequest.NewRequestErrorMessage
							End If
						response.write "</td>"
					response.write "</tr>"
				End If

				response.write "<tr>"
					response.write "<td style='width:150px'>"
						response.write "<span class='th'>Start Date</span>"
					response.write "</td>"
				
					response.write "<td align=left>"
						response.write "<input placeholder='1 Jan 2001' class='evdatepicker' name=fldstrStartDate type=text size=11 maxlength=11 value="""' 
							response.write locobjLeaveRequest.StartDate
						response.write """>"							
					response.write "</td>"
				
				response.write "</tr>"
			
				response.write "<tr>"
					response.write "<td colspan=2>"
						response.write "<div class='th' style='margin-bottom:10px'>Comments</div>"
						response.write "<textarea name=fldstrComments>"
							response.write mHTMLEncode(locobjLeaveRequest.RequestComments)
						response.write "</textarea><br>"												
					response.write "</td>"
				response.write "</tr>"
				
				response.write "<tr>"
					response.write "<td colspan=2>"
						response.write "<br>"
						response.write "<input type=submit value=""Submit ELP Leave Request""><br>"
						response.write "<br>"
					response.write "</td>"
				response.write "</tr>"
			response.write "</table>"		
		response.write "</form>"
		response.write "<br>"
		
			
	end function
	
	function mWriteCalendarHolidayDisplay(ByRef objEEtoView, ByVal teamView)
	
		response.write "<table class='pageContentTable'>"
		response.Write "<tr>"
		response.Write "<td>"
		
			Dim rstLeaveHols
		
			Dim MyCalendar
			Dim cmGetEmployeeHolidayData
			Dim m_cnDB
			Dim viewScope 

			if teamView = True Then
				viewScope = ObjEEtoView.ManagerWWID
			else
				viewScope = ObjEEtoView.WWID
			end if
			
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
			cmGetEmployeeHolidayData.Parameters.Append cmGetEmployeeHolidayData.CreateParameter("@vWWID", adChar, adParamInput, 8, viewScope)
			Set rstLeaveHols = cmGetEmployeeHolidayData.Execute

			' Create the calendar
			Set MyCalendar = New Calendar
			
			' Set the visual properties
			MyCalendar.TitlebarColor = "darkblue" 'Sets the color of the titlebar
			MyCalendar.TitlebarFont = "arial" 'Sets the font face of the titlebar
			MyCalendar.TitlebarFontColor = "white" 'Sets the font color of the titlebar
			MyCalendar.TodayBGColor = "skyblue" 'Sets the highlight color of the current day
			MyCalendar.ShowDateSelect = True 'Toggles the Date Selection form.
			
			'***RUN THROUGH THE FULL LIST OF LEAVE RECORDS
			do while not rstLeaveHols.eof
				dim counter
				dim i
				dim myNewColor
				dim myColorNumber
				
			    myColorNumber=int((999999-100000+1)*rnd+100000)
			    myNewColor = "#" & Cstr(myColorNumber)
				
				'***RUN THROUGH DATES OF LEAVES
				for i = rstLeaveHols.fields.item(2).value to rstLeaveHols.fields.item(3).value
				dim dayOfHol
				dim monOfHol
				dim yearOfHol
				
				dayOfHol = Day(i)
				monOfHol = Month(i)
				yearOfHol = Year(i)
				
				'***CHECK THAT THE DATE IS FOR THE DAY/MONTH/YEAR BEING DISPLAYED AND THEN ADD
				if Year(MyCalendar.GetDate()) = yearOfHol AND Month(MyCalendar.GetDate()) = monOfHol AND not rstLeaveHols.fields.item(5).value = "" AND not  WeekDay(i)=1 AND not WeekDay(i)=7 then
						if not rstLeaveHols.fields.item(6).value = "" then
						
						else
                            
                            if rstLeaveHols.fields.item(7).value then
					            MyCalendar.Days(dayOfHol).AddActivity rstLeaveHols.fields.item(0).value & rstLeaveHols.fields.item(1).value, myNewColor, rstLeaveHols.fields.item(4).value
                            end if
						   
						end if
				end if
				next
				'***MOVE TO NEXT RECORD
				rstLeaveHols.movenext
			loop
			
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
			Dim cmGetPublicHolidayData
			Dim rstPublicHols
			Set cmGetPublicHolidayData = Server.CreateObject("ADODB.Command")
			Set cmGetPublicHolidayData.ActiveConnection =  m_cnDB
			cmGetPublicHolidayData.CommandType = 4
			cmGetPublicHolidayData.CommandText = "dbo.pub_hol_display"
			Set rstPublicHols = cmGetPublicHolidayData.Execute
			
			do while not rstPublicHols.eof
			
				dim dayOfPubHol
				dim monOfPubHol
				dim yearOfPubHol
				
				dayOfPubHol = Day(rstPublicHols.fields.item(1).value)
				monOfPubHol = Month(rstPublicHols.fields.item(1).value)
				yearOfPubHol = Year(rstPublicHols.fields.item(1).value)
				
				'***CHECK THAT THE DATE IS FOR THE DAY/MONTH/YEAR BEING DISPLAYED AND THEN ADD
				if Year(MyCalendar.GetDate()) = yearOfPubHol AND Month(MyCalendar.GetDate()) = monOfPubHol then
						MyCalendar.Days(dayOfPubHol).AddActivity2 rstPublicHols.fields.item(2).value, "black"
				end if
			
				rstPublicHols.movenext
			loop
			
			' Draw the calendar to the browser
			MyCalendar.Draw()
		
		response.Write "</td>"
		response.Write "</tr>"
		response.write "</table>"
		response.write ""
	
	end function
	
	function mWriteAddHolidayForm
	
		if Request.Form("submitHoliday") = "Add Holiday" then
		dim holDate
		dim holName
		dim cmAddHolidays
		dim rsHols
		dim rstHols
		
		holName = Request.Form("nme")
		holDate = Request.Form("fldstrStartDate")
						
		
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
			Set cmAddHolidays = Server.CreateObject("ADODB.Command")
			Set cmAddHolidays.ActiveConnection =  m_cnDB
			cmAddHolidays.CommandType = 4
			cmAddHolidays.CommandText = "dbo.pub_hol_add"
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vMyDate", adDBTimeStamp, adParamInput, , holDate)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vName", adVarChar, adParamInput, 50, holName)
			Set rstHols = cmAddHolidays.Execute
			
		response.write "<table class='pageContentHeader'>"
		response.Write "<tr>"
		response.Write "<td><h3>"
		response.Write "Public Holiday has been added successfully."
		response.Write "</h3></td>"
		response.Write "</tr>"
		response.Write "</table>"
		end if
	
		
		response.write "<form name=frmSQLAdd action=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=phc"" method=post>"

		response.write "<br><table class='pageContentHeader'>"
		response.Write "<tr>"
			response.write "<td><h3>"
					response.write "Add Holidays"
			response.write "</h3></td>"
		response.Write "</tr>"
		response.write "</table>"
		
		response.write "<table class='pageContentTable'>"
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Name of Holiday</span>"
		response.Write "</td>"
		
		'**INPUT FIELD FOR THE NAME OF THE HOLIDAY
		response.Write "<td>"
		response.Write "<input class=txttile type=text name=nme width=150>"
		response.Write "</td>"
		response.Write "</tr>"
		
		response.Write "<tr>"
		response.Write "<td style='width:150px'>"
		response.Write "<span class='th'>Date of Holiday"
		response.Write "<br>"
		response.Write "Format: mon/day/year</span>"
		response.Write "<br>"
		response.Write ""
		
		response.Write "</td>"
		
		'**INPUT FIELD FOR THE DATE OF THE HOLIDAY
		response.write "<td align=left>"
		response.write "<input placeholder='1 Jan 2001' name=fldstrStartDate type=text size=11 maxlength=11 class='evdatepicker'>"
		response.write "</td>"
		
		'**BUTTON CONTROLS OF THE FORM
		response.Write "<tr>"
		response.Write "<td></td><td>"
		response.Write "<input type=reset name=resetButton value=""Reset Form"">"
		response.Write "<input type=submit name=submitHoliday value=""Add Holiday"">"
		response.Write "</td>"
		response.Write "</tr>"
		
		response.Write "</form>"
		response.Write "</table>"
		response.write "<br><br>"
	end function
	
	function mWriteRemoveHolidayForm
	
		Dim m_cnDB
		Dim cmHols
		Dim rstHolidays
		Dim holName
		Dim holDate
		Dim holidayName
		
		'**CHECKS IF FORM HAS BEEN SUBMITTED
		if Request.Form("deleteHoliday") = "Confirm Delete" then
		
		Dim holidayDelName
		Dim holidayDel
		Dim cmDeleteLeaveType
		Dim rstDeleteLeave
		
		'**GETTING THE DATE OF THE HOLIDAY TO BE DELETED		
		holidayDel = Request.Form("holidayDelete")
		
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
			Set cmHols = Server.CreateObject("ADODB.Command")
			Set cmHols.ActiveConnection =  m_cnDB
			cmHols.CommandType = 4
			cmHols.CommandText = "dbo.public_hol_delete"
			cmHols.Parameters.Append cmHols.CreateParameter("@vMyDate", adDBTimeStamp, adParamInput, , holidayDel)			
			Set rstHolidays = cmHols.Execute
			
		response.write "<table class='pageContentHeader'>"
		response.Write "<tr>"
		response.Write "<td><h3>"
		response.Write "Public Holiday has been successfully removed."
		response.Write "</h3></td>"
		response.Write "</tr>"
		response.Write "</table>"
		end if
		
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
			Set cmHols = Server.CreateObject("ADODB.Command")
			Set cmHols.ActiveConnection =  m_cnDB
			cmHols.CommandType = 4
			cmHols.CommandText = "dbo.public_hol_display"
			Set rstHolidays = cmHols.Execute
	
		response.write "<br><table class='pageContentHeader'>"
			response.Write "<tr>"
				response.write "<td><h3>"
					response.write "Delete Holidays"
				response.write "</h3></td>"
			response.Write "</tr>"
		response.Write "</table>"
		
		response.write "<table class='pageContentTable'>"		
			response.write "<form name=frmHolidayRemove action=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=phc"" method=post>"			
			response.Write "<tr>"
				response.Write "<td>"
					response.Write "<select class='basicselect' name=holidayDelete>"
					
						'**CYCLES THROUGH ALL THE DATES CURRENTLY IN THE DATABASE
						do while not rstHolidays.eof
						holDate = rstHolidays.fields.item(0).value
						holName = rstHolidays.fields.item(1).value
						
						'**PRINTS OUT ALL THE DATES AND NAMES OF HOLIDAYS CURRENTLY IN THE DATABASE
						response.Write "<option value=""" & holDate & """>" & holDate & " " & holName & "</option>"
						response.Write "<br>"
						
						rstHolidays.movenext
						loop
					
					response.Write "</select>"
				response.write ""
				response.Write "</td>"
			response.Write "</tr>"
			response.Write "<tr>"
				response.Write "<td>"
					response.Write "<input type=submit name=deleteHoliday value=""Confirm Delete"">"
				response.Write "</td>"
			response.Write "</tr>"
		response.Write "</table>"
		
		response.write "<br><br>"
	
	end function
	
	function mWriteAddLeaveType
	
		'**CHECKS IF THE FORM HAS BEEN SUBMITTED
		if Request.Form("submitLeaveType") = "Add Leave Type" then
		
		dim leaveName
		dim eeRequests
		dim adminRequests
		dim requestBeforeOccrued
		dim minDays
		dim entitlement
		dim daysBeforeStopsLegalAdjAccrual
		dim daysBeforeStopsIsConsecutive
		dim isOtherLeave
		dim cmAddLeaveType
		dim rsHols
		dim rstHols
        dim cmAddHolidays
		
		'**GETTING ALL THE INFORMATION SUBMITTED FROM THE FORM
		leaveName = Request.Form("leaveName")
		eeRequests = Request.Form("userReqs")
		adminRequests = request.Form("adminRequests")
		requestBeforeOccrued = request.Form("requestBeforeOccrued")
		minDays = request.Form("minDays")
		entitlement = request.Form("entitlement")
		daysBeforeStopsLegalAdjAccrual = request.Form("daysBeforeStopsLegalAdjAccrual")
		daysBeforeStopsIsConsecutive = request.Form("daysBeforeStopsIsConsecutive")
		isOtherLeave = request.Form("isOtherLeave")
						
		
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

			Set cmAddHolidays = Server.CreateObject("ADODB.Command")
			Set cmAddHolidays.ActiveConnection =  m_cnDB
			cmAddHolidays.CommandType = 4
			cmAddHolidays.CommandText = "dbo.leave_type_add"
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vName", adVarChar, adParamInput, 50, leaveName)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vEERequests", adInteger, adParamInput, , eeRequests)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vAdminRequests", adInteger, adParamInput, , adminRequests)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vRequestsBeforeOccrued", adInteger, adParamInput, , requestBeforeOccrued)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vMinDays", adVarChar, adParamInput, 50, minDays)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vEntitlement", adVarChar, adParamInput, 50, entitlement)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vDaysBeforeStopsLegalAdjAccrual", adInteger, adParamInput, , daysBeforeStopsLegalAdjAccrual)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vDaysBeforeStopsIsConsecutive", adInteger, adParamInput, , daysBeforeStopsIsConsecutive)
			cmAddHolidays.Parameters.Append cmAddHolidays.CreateParameter("@vIsOtherLeave", adInteger, adParamInput, , isOtherLeave)
			Set rstHols = cmAddHolidays.Execute
			
		response.write "<table class='pageContentHeader'>"
		response.Write "<tr>"
		response.Write "<td>"
		response.Write "Leave Type has been successfully created."
		response.Write "</td>"
		response.Write "</tr>"
		response.Write "</table>"
		end if
		
		'***WARNING
		response.write "<br><table class='pageContentHeader'>"
		response.write "<tr>"
					response.write "<td>"
						response.write "<h3>Warning</h3>"
					response.write "</td>"
				response.write "</tr>"
			response.Write "<tr>"
				response.Write "<td style='padding-left:20px'>"
					response.Write "If you do not understand what a field means then leave the field as is, these are the default settings and will be true for most cases of leave types."
				response.Write "</td>"
			response.Write "</tr>"
		response.Write "</table>"
		
	response.Write ""	
		
		response.write "<br><table class='pageContentHeader'>"
		response.write "<tr>"
				response.write "<td>"
					response.write "<h3>Add Leave Type</h3>"
				response.write "</td>"
			response.write "</tr>"
		response.Write "</table>"
		
		response.write "<table class='pageContentTable'>"
		response.write "<form name=frmLeaveAdd action=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=ltc"" method=post>"
						
		'***NAME OF LEAVE TYPE
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Name of Leave Type</span>"
		response.Write "</td>"
		
		response.Write "<td>"
		response.Write "<input type=text name=leaveName width=150>"
		response.Write "</td>"
		response.Write "</tr>"
		
		'***USER REQUESTS
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>User Requests</span>"
		response.Write "</td>"
		
		response.write "<td align=left>"
		response.write "<input type=radio name=userReqs value=1 checked> True"
		response.Write "<br>"
		response.write "<input type=radio name=userReqs value=0> False"
		response.write "</td>"
		response.Write "<tr>"
		
		'***ADMIN REQUESTS
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Admin Requests</span>"
		response.Write "</td>"
		
		response.write "<td align=left>"
		response.write "<input type=radio name=adminRequests value=1 checked> True"
		response.Write "<br>"
		response.write "<input type=radio name=adminRequests value=0> False"
		response.write "</td>"
		response.Write "<tr>"
		
		'***REQUESTS BEFORE OCCRUED
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Requests Before Accrued</span>"
		response.Write "</td>"
		
		response.write "<td align=left>"
		response.write "<input type=radio name=requestBeforeOccrued value=1> True"
		response.Write "<br>"
		response.write "<input type=radio name=requestBeforeOccrued value=0 checked> False"
		response.write "</td>"
		response.Write "<tr>"
		
		'***MINIMUM DAYS
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Minimum Days</span>"
		response.Write "</td>"
		
		response.Write "<td>"
		response.Write "<input class=txttile type=text name=minDays width=150 value=0.5>"
		response.Write "</td>"
		response.Write "</tr>"
		
		'***ENTITLEMENT
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Entitlement</span>"
		response.Write "</td>"
		
		response.Write "<td>"
		response.Write "<input class=txttile type=text name=entitlement width=150 value=0>"
		response.Write "</td>"
		response.Write "</tr>"
		
		'***DAYS BEFORE STOPS LEGAL ADJ ACCRUAL
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Days before stops Legal Adj Accrual</span>"
		response.Write "</td>"
		
		response.Write "<td>"
		response.Write "<input class=txttile type=text name=daysBeforeStopsLegalAdjAccrual width=150 value=0>"
		response.Write "</td>"
		response.Write "</tr>"
		
		'***DAYS BEFORE STOPS IS CONSECUTIVE
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Days before stops is Consecutive</span>"
		response.Write "</td>"
		
		response.write "<td align=left>"
		response.write "<input type=radio name=daysBeforeStopsIsConsecutive value=1> True"
		response.Write "<br>"
		response.write "<input type=radio name=daysBeforeStopsIsConsecutive value=0 checked> False"
		response.write "</td>"
		response.Write "<tr>"
		
		'***IS OTHER LEAVE
		response.Write "<tr>"
		response.Write "<td style='width:350px'>"
		response.Write "<span class='th'>Is Other Leave</span>"
		response.Write "</td>"
		
		response.write "<td align=left>"
		response.write "<input type=radio name=isOtherLeave value=1 checked> True"
		response.Write "<br>"
		response.write "<input type=radio name=isOtherLeave value=0> False"
		response.write "</td>"
		response.Write "<tr>"
		
		'***SUBMISSION AND RESET BUTTONS
		response.Write "<td></td><td>"
		response.Write "<input type=reset name=resetButton value=""Reset Form"">"
		response.Write "<input type=submit name=submitLeaveType value=""Add Leave Type"">"
		response.Write "</td>"
		response.Write "</tr>"
		
		response.Write "</form>"
		response.Write "</table>"
		response.write ""
	
	end function
	
	function mWriteRemoveLeaveType
	
		Dim m_cnDB
		Dim cmLeaves
		Dim rstLeaveTypes
		Dim leaveName
			
		'**CHECKS IF THE FORM HAS BEEN SUBMITTED
		if Request.Form("deleteLeaveType") = "Confirm Delete" then
			Dim leaveTypeName
			Dim cmDeleteLeaveType
			Dim rstDeleteLeave
			
			'**GETS NAME OF THE LEAVE TYPE TO BE DELETED FROM THE FORM
			leaveTypeName = Request.Form("leaveDelete")
			
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
			Set cmDeleteLeaveType = Server.CreateObject("ADODB.Command")
			Set cmDeleteLeaveType.ActiveConnection =  m_cnDB
			cmDeleteLeaveType.CommandType = 4 

            cmDeleteLeaveType.CommandText = "dbo.leave_type_delete_check"
			cmDeleteLeaveType.Parameters.Append cmDeleteLeaveType.CreateParameter("@vName", adVarChar, adParamInput, 50, leaveTypeName)
			Set rstDeleteLeave = cmDeleteLeaveType.Execute

            if rstDeleteLeave("amount") = 0 then
			    cmDeleteLeaveType.CommandText = "dbo.leave_type_delete"
			    cmDeleteLeaveType.Parameters.Append cmDeleteLeaveType.CreateParameter("@vName", adVarChar, adParamInput, 50, leaveTypeName)
			    Set rstDeleteLeave = cmDeleteLeaveType.Execute
			
			    response.write "<table class='pageContentHeader'>"
			    response.Write "<tr>"
			    response.Write "<td style='text-align:center'>"
			    response.Write "Leave Type has been successfully removed."
			    response.Write "</td>"
			    response.Write "</tr>"
			    response.Write "</table>"
            else 
                 mWriteGeneralError "Sorry - you cannot delete a leave type associated with existing leave requests.", False
            end if
		end if

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
		Set cmLeaves = Server.CreateObject("ADODB.Command")
		Set cmLeaves.ActiveConnection =  m_cnDB
		cmLeaves.CommandType = 4
		cmLeaves.CommandText = "dbo.leave_type_display"
		Set rstLeaveTypes = cmLeaves.Execute


		
		response.write "<form name=frmLeaveRemove action=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=ltc"" method=post>"
		
		'***HEADER
		response.write "<table class='pageContentHeader'>"
		response.Write "<tr>"
		response.write "<td>"
		response.write "<h3>Remove Leave Type</h3>"
		response.write "</td>"
		response.Write "</tr>"
		response.write "</table>"
		
		response.write "<table class='pageContentTable'>"
		response.Write "<tr>"
		response.Write "<td>"
		
		response.Write "<select class='basicselect' name=leaveDelete>"
		
		'**CYCLES THROUGH ALL LEAVE TYPES CURRENTLY STORED IN THE DATABASE
		do while not rstLeaveTypes.eof
			leaveName = rstLeaveTypes.fields.item(0).value
			response.Write "<option>" & leaveName & "</option>"
			response.Write "<br>"
			
			rstLeaveTypes.movenext
		loop
		
		response.Write "</select>"
		
		response.Write "</td>"
		response.Write "</tr>"
		response.Write "<tr>"
		response.Write "<td>"
		response.write ""
		response.Write "<input type=submit name=deleteLeaveType value=""Confirm Delete"">"
		response.Write "</td>"
		response.Write "</tr>"
		response.Write "</form>"
		response.Write "</table>"
		response.write "<br><br>"
		
	end function
	
	function mWriteSqlLinks
		response.write "<br><table class='pageContentHeader'>"
		
	    '***HEADER
		response.Write "<tr>"
		response.write "<td>"
		response.write "<h3>Form Changes Available</h3>"
		response.write "</td>"
		response.Write "</tr>"
		response.Write "</table>"
		
		response.write "<br><table class='pageContentTable'>"
		response.Write "<tr>"
		'**WRITES THE LINKS TO THE DIFFERENT PAGES
		response.Write "<td >"
			response.write "<a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=ltc"" title=""Add/Remove Leave Types."">"
			response.Write "Leave Types Form"
			response.Write "</a>"
			response.write "<br><br>"
			response.write "<a href=""" & CONST_APPLICATION_PATH & "/adminhome.asp?m=phc"" title=""Add/Remove Public Holidays."">"
			response.Write "Public Holidays Form"
			response.Write "</a>"
		response.Write "</td>"
		response.Write "</tr>"
		response.Write "</table>"
		response.write "<br><br>"
	end function

%>