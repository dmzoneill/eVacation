<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	strCurrentPageName = "ELP Summary"
	
	Dim locobjELPLeavePeriod
	Dim loclngResult
	Dim locblnDisplaySummary
	Dim locblnDisplayCancelRequestForm
	Dim locobjELPLeaveRequest
	Dim locblnValidateRequest
	Dim loclngSaveResult
	
	mInitialiseCurrentUser
	
	If not objCurrentUser.IsEELeaveTracked then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If

	locblnDisplaySummary = True

	
	Set locobjELPLeaveRequest = new cObjLeavePeriod
	
	If request.form("formname") = "frmRequestELPLeave" then
		locblnValidateRequest = True
		locobjELPLeaveRequest.LoadNewRequestFromForm
	Else
		locblnValidateRequest = False
	End If
		
	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

	if locblnValidateRequest then
		if locobjELPLeaveRequest.FormIsValid then
			if locobjELPLeaveRequest.NewRequestIsValid then
				
				loclngSaveResult = locobjELPLeaveRequest.Save
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
	If strMode = "cl" then
		Set locobjELPLeavePeriod = new cObjLeavePeriod
		locobjELPLeavePeriod.ID = lngItemID
		If not locobjELPLeavePeriod.IsActive then
			mWriteGeneralError "Sorry - the leave period you are attempting to cancel is not active.", False
		Else
			If locobjELPLeavePeriod.CancelRequiresApproval then
				locblnDisplayCancelRequestForm = True
				locblnDisplaySummary = False
				If request("formname") = "frmCancelLeavePeriod" then
					locblnValidateRequest = True
					locobjELPLeavePeriod.LoadCancelRequestFromForm
					If locobjELPLeavePeriod.CancelRequestFormIsValid then
						loclngResult = locobjELPLeavePeriod.Cancel
						If loclngResult = 1 then
							mWriteGeneralError "Sorry - an error has occurred while attempting to raise the cancellation request.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
						Else
							response.write "<center>"
								response.write "<br>"
								response.write "<b>The cancellation request has been raised successfully.</b><br>"
								response.write "<br>"
							response.write "</center>"
							Set locobjELPLeavePeriod = nothing
							locblnDisplayCancelRequestForm = False
							locblnDisplaySummary = True
						End If
					End If
				Else
					locblnValidateRequest = False
				End If
			else
				loclngResult = locobjELPLeavePeriod.Cancel
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
			End If
		End If
	End If

	'**** INITIALISE CURRENT USER ****
	mInitialiseEEtoView
	
	Set objEEtoView = objCurrentUser
		
	'*** Display Cancellation Request form
	If locblnDisplayCancelRequestForm then
		mWriteCancelLeavePeriodRequestForm objEEtoView, locobjELPLeavePeriod, locblnValidateRequest
	End If
	
	If locblnDisplaySummary then
		mWriteViewingEmployee objEEtoView
		
		
		If objEEtoView.AnnualVacation.HasMaturedELP OR _
		   objEEtoView.AnnualVacation.ELPUsed.IsUsed OR _
		   objEEtoView.AnnualVacation.HasActiveELP then
			
			'***Mature ELP
			If objEEtoView.AnnualVacation.HasMaturedELP then
				mWriteMaturedELPDetails objEEtoView.AnnualVacation.ELPMatured, false
				If not objEEtoView.AnnualVacation.ELPMatured.IsExpired then
					mWriteELPLeaveRequestForm objEEtoView, locobjELPLeaveRequest, locblnValidateRequest
				End If
			End If
			
			'**ELP Vacation Details
			If objEEtoView.AnnualVacation.ELPUsed.IsUsed then
				mWriteELPVacationDetails objEEtoView, objEEtoView.AnnualVacation.ELPUsed, false
			End If
			
			'***Active ELP
			If objEEtoView.AnnualVacation.HasActiveELP then
				mWriteELPDetails objEEtoView.AnnualVacation.ELPActive, False
			End If
		Else
			response.write "<br><table class='pageContentHeader'><tbody><tr><td><h3>Extended Leave Program</h3></td></tr>"
						
			response.write "<tr>"
				response.write "<td style='text-align:center'>"
					response.write "You are not shown as being registered in the Extended Leave Program.<br>"
					response.write "Contact e-Vacation Administrator to register"
				response.write "</td>"
			response.write "</tr>"
			response.write "</table><br><br>"	
		End If	
	End If	
	
	mWritePageFooter
	
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
