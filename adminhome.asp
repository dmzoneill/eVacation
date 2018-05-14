<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	Dim locblnUserSelected
	Dim locstrAdminMenuOption
	Dim locblnValidateUser
	Dim locblnValidateProfile
	Dim locblnValidateCarryOver
	Dim locblnValidateLeavePeriod
	Dim locblnValidateActivateELP
	Dim locblnValidateQuery
	Dim locobjUser
	Dim locobjCarryOver
	Dim locobjLeavePeriod
	Dim locobjLeavePeriodCollection
	Dim locobjActiveELP
	Dim locobjELPRelief
	Dim lngViewYear
	Dim locLeaveRequest
	
	Dim loclngSaveResult
	
	strCurrentPageName = "Admin Menu"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView
	
	
	'**** CHECK THAT THE CURRENT USER HAS ADMINISTRATOR RIGHTS.
	If not objCurrentUser.IsAdmin then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If

	Set locobjUser = new cObjUser
	locobjUser.WWID = request("ee")
      '   locobjUser.WWID = 10690153
	If locobjUser.WWID = "" then
		locblnUserSelected = False
	Elseif not locobjUser.AdminSelectUserFormIsValid then
		locblnValidateUser = True
		locblnUserSelected = False
	Else
		locblnUserSelected = True
	End If

	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName
	
	if (strMode = "rp") or (strMode ="pr") or (strMode = "hr") then
		locstrAdminMenuOption = "Reports"
		locblnUserSelected = True
		
		elseif (strMode = "dq") then '[MFILLAST 08-2006] added database query
		locstrAdminMenuOption = "Database Query"
		locblnUserSelected = True
		
		elseif (strMode = "sfm") then
		locstrAdminMenuOption = "SQL Forms"
		locblnUserSelected = True
		
		elseif (strMode = "ltc") then
		locstrAdminMenuOption = "Leave Type Changes"
		locblnUserSelected = True
		
		
		elseif (strMode = "phc") then
		locstrAdminMenuOption = "Public Holiday Changes"
		locblnUserSelected = True
	

	elseif not locblnUserSelected then
		locstrAdminMenuOption = "Select User"
		Set objEEtoView = nothing
		Set objEEtoView = new cObjUser
	else
		Select Case strMode
			Case "co"
				locstrAdminMenuOption = "Carry Over"
				Set locobjCarryOver = new cObjCarryOver
				If request("formname") = "frmCarryOver" then
					locobjCarryOver.LoadAdminCarryOverFromForm
					If not locobjCarryOver.AdminCarryOverFormIsValid then
						locblnValidateCarryOver = True
					Else
						loclngSaveResult = locobjCarryOver.Save
						If loclngSaveResult = 0 then
							mWriteGeneralError "Sorry - an error has occurred while attempting to save the carry over entry.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
						else
							response.write "<center>"
								response.write "<br>"
								response.write "<b>The carry over entry has been updated successfully.</b><br>"
								response.write "<br>"
							response.write "</center>"
							locblnValidateCarryOver = False
						end if
					End If
				End If
			Case "lv"
				locstrAdminMenuOption = "Leave"
				Set locobjLeavePeriod = new cObjLeavePeriod
				locobjLeavePeriod.IsAdminUpdating = True
				If request("formname") = "frmLeavePeriod" then
					if request("btnSubmit") = "Save" then
						locobjLeavePeriod.LoadAdminLeavePeriodFromForm
						If not locobjLeavePeriod.AdminLeavePeriodFormIsValid then
							locblnValidateLeavePeriod = True
						ElseIf not locobjLeavePeriod.NewAdminRequestIsValid then
							locblnValidateLeavePeriod = True
						Else
							loclngSaveResult = locobjLeavePeriod.Save
							If loclngSaveResult = 0 then
								mWriteGeneralError "Sorry - an error has occurred while attempting to save the leave entry.<br>" & _
									"There may be a problem with the network, or the database server may be experiencing difficulties.", False
							else
								response.write "<center>"
									response.write "<br>"
									response.write "<b>The leave entry has been updated successfully.</b><br>"
									response.write "<br>"
								response.write "</center>"
								locblnValidateLeavePeriod = False
								Set locobjLeavePeriod = nothing
								Set locobjLeavePeriod = new cObjLeavePeriod
							end if
						End If
					elseif request("btnSubmit") = "Delete" then
						locobjLeavePeriod.LoadAdminLeavePeriodFromForm
						loclngSaveResult = locobjLeavePeriod.Delete
						If loclngSaveResult <> 0 then
							mWriteGeneralError "Sorry - an error has occurred while attempting to delete the leave entry.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
						else
							response.write "<center>"
								response.write "<br>"
								response.write "<b>The leave entry has been deleted successfully.</b><br>"
								response.write "<br>"
							response.write "</center>"
							locblnValidateLeavePeriod = False
							Set locobjLeavePeriod = nothing
							Set locobjLeavePeriod = new cObjLeavePeriod
						end if
					end if
				ElseIf lngItemID <> 0 then
					locobjLeavePeriod.ID = lngItemID
					If locobjLeavePeriod.EE.WWID <> locobjUser.WWID then
						Set locobjLeavePeriod = nothing
						Set locobjLeavePeriod = new cObjLeavePeriod
						lngItemID = 0
					End If
				End If

			Case "eelp"
								
				locstrAdminMenuOption = "ELP"
				
				If lngItemID = 0 then
					mWriteGeneralError "Sorry - the ELP instance you are trying to edit was not recognised.", 0
				Else
					Set locobjActiveELP = new cObjELPInstance
					locobjActiveELP.ELPID = lngItemID
					
					
					If request("formname") = "frmActivateELP" then
						locobjActiveELP.LoadAdminActivateELPFromForm
						If not locobjActiveELP.AdminActivateELPFormIsValid then
							locblnValidateActivateELP = True
						Else
							loclngSaveResult = locobjActiveELP.Save
							If loclngSaveResult = 0 then
								mWriteGeneralError "Sorry - an error has occurred while attempting to process your request to activate ELP.<br>" & _
									"There may be a problem with the network, or the database server may be experiencing difficulties.", False
							else
								response.write "<center>"
									response.write "<br>"
									response.write "<b>ELP has been updated successfully.</b><br>"
									response.write "<br>"
								response.write "</center>"
								locblnValidateActivateELP = False
								'*** Reload the user object to make sure that the ELP Instance is refreshed up.
								Set locobjUser = nothing
								Set locobjUser = new cObjUser
								locobjUser.WWID = request("ee")
								
								Set locobjActiveELP = locobjUser.AnnualVacation.ELPActive
								
								strMode = "elp"
							end if
						End If
					End If
				End If

			Case "elp"
				
				locstrAdminMenuOption = "ELP"
				
				Set locobjELPRelief = new cObjELPRelief
				
				If locobjUser.AnnualVacation.HasActiveELP then
					 Set locobjActiveELP = locobjUser.AnnualVacation.ELPActive
								
					If request("formname") = "frmEditELPRelief" then
					
						locobjELPRelief.AdminELPReliefLoadFromForm
						loclngSaveResult = locobjELPRelief.Save
						If loclngSaveResult = 0 then
							mWriteGeneralError "Sorry - an error has occurred while attempting to update ELP Reliefs.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
						else
								response.write "<center>"
									response.write "<br>"
									response.write "<b>ELP Reliefs have been updated successfully.</b><br>"
									response.write "<br>"
								response.write "</center>"
						end if
						'*** Reload the user object to make sure that the updated ELP Reliefs are picked up.
						Set locobjUser = nothing
						Set locobjUser = new cObjUser
						locobjUser.WWID = request("ee")
						Set locobjActiveELP = nothing
						Set locobjActiveELP = locobjUser.AnnualVacation.ELPActive
					End If

				Else
					
					Set locobjActiveELP = new cObjELPInstance
					Set locobjActiveELP.EE = locobjUser
					
					If request("formname") = "frmActivateELP" then
						locobjActiveELP.LoadAdminActivateELPFromForm
						If not locobjActiveELP.AdminActivateELPFormIsValid then
							locblnValidateActivateELP = True
						Else
							loclngSaveResult = locobjActiveELP.Save
							If loclngSaveResult = 0 then
								mWriteGeneralError "Sorry - an error has occurred while attempting to process your request to activate ELP.<br>" & _
									"There may be a problem with the network, or the database server may be experiencing difficulties.", False
							else
								response.write "<center>"
									response.write "<br>"
									response.write "<b>ELP has been activated successfully.</b><br>"
									response.write "<br>"
								response.write "</center>"
								locblnValidateActivateELP = False
								'*** Reload the user object to make sure that the new ELP Instance is picked up.
								Set locobjUser = nothing
								Set locobjUser = new cObjUser
								locobjUser.WWID = request("ee")
							end if
						End If
					End If
						
				End If
				
			Case Else
				locstrAdminMenuOption = "User Profile"
				If request("formname") = "frmUserProfile" then
					locobjUser.LoadAdminUserProfileFromForm
					If not locobjUser.AdminUserProfileFormIsValid then
						locblnValidateProfile = True
					Else
						loclngSaveResult = locobjUser.Save
						If loclngSaveResult = 0 then
							mWriteGeneralError "Sorry - an error has occurred while attempting to save the user profile.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
						else
							response.write "<center>"
								response.write "<br>"
								response.write "<b>The user's profile has been updated successfully.</b><br>"
								response.write "<br>"
							response.write "</center>"
							locblnValidateProfile = False
						end if
					End If
				End If
		End Select
	end if
	
	mWriteAdminMenu objEEtoView, locstrAdminMenuOption
	
	If not locblnUserSelected then
		mWriteAdminSelectUserForm locobjUser, locblnValidateUser
	Else
		Select Case strMode
			Case "co"
				mWriteAdminCarryOver locobjUser, locobjCarryOver, locblnValidateCarryOver
			Case "lv"
				lngViewYear = mGetSafeLongInteger(request("lngYear"),0)
				
				'*******TESTING
				'Response.Write lngViewYear & "<BR>"
				
				If lngViewYear <> 0 and lngViewYear >= CONST_FIRST_YEAR_SYSTEM_ACTIVE then
					locobjUser.YearToView = lngViewYear
				
					
				End If
				
				
				mWriteAdminLeavePeriod locobjUser, locobjLeavePeriod, locblnValidateLeavePeriod
				
				mWriteLeaveSummary locobjUser, CONST_APPLICATION_PATH & "/adminhome.asp"
				
				mWriteLeaveRequests locobjUser.LeaveRequests, true
				
					If objEEtoView.AnnualVacation.HasActiveELP then
						mWriteELPDetails objEEtoView.AnnualVacation.ELPActive, False
					End If
					If objEEtoView.AnnualVacation.HasMaturedELP then
						mWriteELPDetails objEEtoView.AnnualVacation.ELPMatured, False
					End If
			Case "eelp"
				mWriteELPEditForm locobjActiveELP
			Case "elp"
				mWriteAdminELP locobjUser, locobjActiveELP, locobjELPRelief, locblnValidateActivateELP
			Case "rp"
				mWriteReportFormBase
			case "pr"
				mWritePayrollReportForm
			case "hr"
				mWriteHRReportForm
			case "dq" 
				mWriteDatabaseQueryForm
			Case "sfm"
				mWriteSqlLinks
			Case "ltc"
				mWriteAddLeaveType
				mWriteRemoveLeaveType
			Case "phc"
				mWriteAddHolidayForm
				mWriteRemoveHolidayForm
				mWritePublicHolidays  
			Case Else
				mDebugPrint "Calling mWriteAdminUserProfileForm objUser, '" & locblnValidateProfile & "'<br>"
				mDebugPrint "User.WWID = '" & locobjUser.WWID & "'<br>"
				mWriteAdminUserProfileForm locobjUser, locblnValidateProfile
		End Select
	End If
	
	mWritePageFooter

	Set locobjELPRelief = nothing
	Set locobjActiveELP = nothing
	Set locobjLeavePeriodCollection = nothing
	Set locobjCarryOver = nothing
	Set locobjLeavePeriod = nothing
	Set locobjUser = nothing
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
