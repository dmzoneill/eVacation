<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%

	Dim lngViewYear
	Dim locobjLeavePeriod
	Dim loclngResult
	Dim locblnDisplaySummary
	Dim locblnDisplayCancelRequestForm
	Dim locblnValidateRequest

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser

	strCurrentPageName = "Home Page"
	Set objEEtoView = objCurrentUser

	locblnDisplaySummary = True

	lngViewYear = mGetSafeLongInteger(request("lngYear"),0)
	If lngViewYear <> 0 and lngViewYear >= CONST_FIRST_YEAR_SYSTEM_ACTIVE then
		objEEtoView.YearToView = lngViewYear
	End If
		
	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

	response.Write "<br class=small>"
	
	' Cancel Mode
	If strMode = "cl" then
		Set locobjLeavePeriod = new cObjLeavePeriod
		locobjLeavePeriod.ID = lngItemID
		If not locobjLeavePeriod.IsActive then
			mWriteGeneralError "Sorry - the leave period you are attempting to cancel is not active.", False
		Else
			If locobjLeavePeriod.CancelRequiresApproval then
				locblnDisplayCancelRequestForm = True
				locblnDisplaySummary = False
				If request("formname") = "frmCancelLeavePeriod" then
					locblnValidateRequest = True
					locobjLeavePeriod.LoadCancelRequestFromForm
					If locobjLeavePeriod.CancelRequestFormIsValid then
						loclngResult = locobjLeavePeriod.Cancel
						If loclngResult = 1 then
							mWriteGeneralError "Sorry - an error has occurred while attempting to raise the cancellation request.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
						Else
							response.write "<center>"
								response.write "<br>"
								response.write "<b>The cancellation request has been raised successfully.</b><br>"
								response.write "<br>"
							response.write "</center>"
							Set locobjLeavePeriod = nothing
							locblnDisplayCancelRequestForm = False
							locblnDisplaySummary = True
						End If
					End If
				Else
					locblnValidateRequest = False
				End If
			else
				loclngResult = locobjLeavePeriod.Cancel
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

	If locblnDisplayCancelRequestForm then
		mWriteCancelLeavePeriodRequestForm objEEtoView, locobjLeavePeriod, locblnValidateRequest
	End If

	' [MOF 11/12/08] Added "Confirm" Mode 
	If strMode = "cf" then
		' Load the Leave Period
		Set locobjLeavePeriod = new cObjLeavePeriod
		locobjLeavePeriod.ID = lngItemID
		
		If not locobjLeavePeriod.IsActive then
			mWriteGeneralError "Sorry - the leave period you are attempting to confirm is not active.", False
		Else
			loclngResult = locobjLeavePeriod.Confirm
			If loclngResult = 1 then
				mWriteGeneralError "Sorry - an error has occurred while attempting to confirm the leave period.<br>" & _
								"There may be a problem with the network, or the database server may be experiencing difficulties.", False
			Else
				' Print confirm message
				response.write "<center>"
					response.write "<br>"
					response.write "<b>The leave request has been confirmed successfully.<br>"
					response.write "<br>"
					response.write "Thank you.</b><br>" 
					response.write "<br>"
				response.write "</center>"
			End If
		End If
	End If
	' [/MOF]
		
	If locblnDisplaySummary then
		mWriteViewingEmployee objEEtoView
		mWriteLeaveSummary objEEtoView, CONST_APPLICATION_PATH & "/default.asp"
		mWriteCompTimeDetails objEEtoView.AnnualVacation.CompTime, false
		mWriteLeaveRequests objEEtoView.LeaveRequests, false
		If objEEtoView.AnnualVacation.HasActiveELP then
			mWriteELPDetails objEEtoView.AnnualVacation.ELPActive, False
		End If
		If objEEtoView.AnnualVacation.HasMaturedELP then
			mWriteELPDetails objEEtoView.AnnualVacation.ELPMatured, False
		End If
	End If
	
	mWritePageFooter
	
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
