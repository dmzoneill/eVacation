<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	strCurrentPageName = "Approve Requests"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser

	Dim locobjLeaveRequest
	Dim locblnValidateApproval
	Dim locblnDisplayUserApprovals
	Dim locblnDisplayApproveLeaveRequest
	Dim loclngSaveResult
	Dim locblnPageHeaderWritten
	Dim locstrApprovalAction
	
	locblnPageHeaderWritten = False

	Select Case strMode
		Case CONST_MODE_APPROVE_REQUESTS_VIEW_REQUEST
			locblnDisplayApproveLeaveRequest = True
			locblnDisplayUserApprovals = False

			Set locobjLeaveRequest = new cObjLeavePeriod

			locobjLeaveRequest.ID = lngItemID

			if locobjLeaveRequest.AwaitingApproval then
				If request.form("formname") = "frmLeaveApproval" then
					locobjLeaveRequest.LoadLeaveApprovalFromForm
					locblnValidateApproval = True
				Else
					locblnValidateApproval = False
				End If
	
				'*** TEST TO MAKE SURE THE SPECIFIED LEAVE REQUEST WAS FOUND ***
				if locobjLeaveRequest.EE.WWID = "" then
					mWriteHMTLTop strCurrentPageName
					mWriteNavBar strCurrentPageName
					locblnPageHeaderWritten = True
					mWriteGeneralError "Sorry - the leave request selected was not found.", False
					locblnDisplayApproveLeaveRequest = False
					locblnDisplayUserApprovals = True
	
				'*** TEST TO MAKE SURE THAT THE CURRENT USER IS A VALID APPROVER OF THIS LEAVE REQUEST ***
				elseif not locobjLeaveRequest.IsValidApprover(objCurrentUser) then
					mWriteHMTLTop strCurrentPageName
					mWriteNavBar strCurrentPageName
					locblnPageHeaderWritten = True
					mWriteGeneralError "Sorry - you do not have the authority to approve or reject the leave request specified.", False
					locblnDisplayApproveLeaveRequest = False
					locblnDisplayUserApprovals = True
	
				elseif locblnValidateApproval then
					if locobjLeaveRequest.ApprovalFormIsValid then
						loclngSaveResult = locobjLeaveRequest.Save
						If loclngSaveResult = 0 then
							mWriteHMTLTop strCurrentPageName
							mWriteNavBar strCurrentPageName
							locblnPageHeaderWritten = True
							mWriteGeneralError "Sorry - an error has occurred while attempting to action your request.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
						else
							Select Case locobjLeaveRequest.Status
								Case CONST_LEAVE_PERIOD_STATUS_APPROVED, CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED
									locstrApprovalAction = "approved"
								Case CONST_LEAVE_PERIOD_STATUS_REJECTED, CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED
									locstrApprovalAction = "rejected"
							End Select
							
							Set locobjLeaveRequest = nothing
							Set locobjLeaveRequest = new cObjLeavePeriod
							locblnDisplayApproveLeaveRequest = False
	
							objCurrentUser.RefreshApprovals
	
							mWriteHMTLTop strCurrentPageName
							mWriteNavBar strCurrentPageName
							locblnPageHeaderWritten = True
	
							response.write "<center>"
								response.write "<br>"
								response.write "<span class=txtextlarge>Thank You.</span><br><br>"
								response.write "<b><span class=txtlarge>You have successfully "
								response.write locstrApprovalAction
								response.write " the request.</span></b><br>"
								response.write "<br>"
							response.write "</center>"
	
							locblnDisplayUserApprovals = True
						end if
					end if
				end if
			else
				mWriteHMTLTop strCurrentPageName
				mWriteNavBar strCurrentPageName
				locblnPageHeaderWritten = True
				mWriteGeneralError "Sorry - item specified has already been actioned.", False
				locblnDisplayApproveLeaveRequest = False
				locblnDisplayUserApprovals = True
			end if	

			If locblnDisplayApproveLeaveRequest then
				If not locblnPageHeaderWritten then
					mWriteHMTLTop strCurrentPageName
					mWriteNavBar strCurrentPageName
					locblnPageHeaderWritten = True
				End If
				mWriteApproveLeaveRequest locobjLeaveRequest, locblnValidateApproval
			End If

			Set locobjLeaveRequest = nothing 

		Case Else
			locblnDisplayUserApprovals = True

	End Select

	If not locblnPageHeaderWritten then
		mWriteHMTLTop strCurrentPageName
		mWriteNavBar strCurrentPageName
		locblnPageHeaderWritten = True
	End If

	if locblnDisplayUserApprovals then
		mWriteApproveRequests objCurrentUser
	end if
	
	mWritePageFooter
	
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
