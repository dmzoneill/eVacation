<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	Dim lngViewYear
	Dim locobjLeavePeriod
	Dim loclngResult
	Dim locblnDisplaySummary
	Dim locblnDisplayCancelRequestForm
	Dim locblnValidateRequest
	Dim locblnShowConfirmedMsg
	Dim locblnShowCancelRaisedMsg
	Dim locblnShowCancelSuccessMsg
	Dim locErrorMsg
	locErrorMsg = ""
	locblnDisplaySummary = True

    mInitialiseCurrentUser

	'Added Confirm Leave logic, confirm/error messages are printed below
	locblnShowConfirmedMsg = false
	If strMode = "cf" and Request.Form("formname") = "" then
		' Load the Leave Period
		Set locobjLeavePeriod = new cObjLeavePeriod
		locobjLeavePeriod.ID = lngItemID

        if locobjLeavePeriod.IsValidApprover(objCurrentUser) or objCurrentUser.WWID = locobjLeavePeriod.EE.WWID then
		
		    ' Check if the leave period can be confirmed
		    If not locobjLeavePeriod.IsActive Then
			    locErrorMsg = "Sorry - the leave period you are attempting to confirm as taken is not active."
		    Elseif locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED or _
		           locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED Then
			    locErrorMsg = "Sorry - the leave period you are attempting to confirm as taken has been cancelled."
		    Else
			    loclngResult = locobjLeavePeriod.Confirm
			    If loclngResult = 1 then
				    locErrorMsg = "Sorry - an error has occurred while attempting to confirm the leave period as taken.<br>" & _
							      "There may be a problem with the network, or the database server may be experiencing difficulties."
			    Else
				    ' Print confirm message
				    locblnShowConfirmedMsg = true
			    End If
		    End If
        End if
	End If

	' [MOF 16/12/08] Moved "Cancel Leave" Logic above InitialiseUser so that Leave Requests will be up to date when displayed
	locblnShowCancelRaisedMsg = False
	locblnShowCancelSuccessMsg = False
	If strMode = "cl" then
		Set locobjLeavePeriod = new cObjLeavePeriod
		locobjLeavePeriod.ID = lngItemID

        if locobjLeavePeriod.IsValidApprover(objCurrentUser) or objCurrentUser.WWID = locobjLeavePeriod.EE.WWID then
		   
            if not locobjLeavePeriod.IsActive or locobjLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CONFIRMED then
			    locErrorMsg = "Sorry - the leave period you are attempting to cancel is not active."
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
							    locErrorMsg = "Sorry - an error has occurred while attempting to raise the cancellation request.<br>" & _
								    "There may be a problem with the network, or the database server may be experiencing difficulties."
						    Else
							    locblnShowCancelRaisedMsg = True
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
					    locErrorMsg = "Sorry - an error has occurred while attempting to cancel the leave period.<br>" & _
								    "There may be a problem with the network, or the database server may be experiencing difficulties."
				    Else
					    locblnShowCancelSuccessMsg = True
				    End If
			    End If
		    End If
        End if
	End If
	
	
	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	
	Select Case strMode
		Case "el"
			strCurrentPageName = "Team Leave Summary"
			mInitialiseEEtoView
			
			Dim rstMansMan
			Dim cmGetManagersMan
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
								Set cmGetManagersMan = Server.CreateObject("ADODB.Command")
								Set cmGetManagersMan.ActiveConnection =  m_cnDB
								cmGetManagersMan.CommandType = 4
								cmGetManagersMan.CommandText = "dbo.mans_man"
								cmGetManagersMan.Parameters.Append cmGetManagersMan.CreateParameter("@vWWID", adChar, adParamInput, 8, objEEtoView.Manager.WWID)
								Set rstMansMan = cmGetManagersMan.Execute
			
			If objEEtoView.Manager.WWID <> objCurrentUser.WWID AND rstMansMan.fields.item(0).value <> objCurrentUser.WWID then
				mCloseApplication
				response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
			End If
		Case Else
			strCurrentPageName = "Leave Summary"
			Set objEEtoView = objCurrentUser
	End Select

	lngViewYear = mGetSafeLongInteger(request("lngYear"),0)
	If lngViewYear <> 0 and lngViewYear >= CONST_FIRST_YEAR_SYSTEM_ACTIVE then
		objEEtoView.YearToView = lngViewYear
	End If
	
	
	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

	' [MOF 11/17/08] Added Confirm Leave Message (Logic above)
	' NOTE: Had to use confirm message flag so that logic was performed before user was loaded
	If locblnShowConfirmedMsg then
		response.write "<div class='feedbackbox'>"
			response.write "The leave request has been confirmed as taken successfully."
		response.write "</div>"
	End If
	If locblnShowCancelRaisedMsg Then
        response.write "<div class='feedbackbox'>"
			response.write "The cancellation request has been raised successfully."
		response.write "</div>"
	End If
	If locblnShowCancelSuccessMsg Then
		response.write "<div class='feedbackbox'>"
			response.write "The leave period has been cancelled successfully."
		response.write "</div>"
	End If  
	
	' [MOF 16/12/08] Moved logic to top of page and kept error messages displaying here
	if "" <> locErrorMsg then
	    mWriteGeneralError locErrorMsg, False
	end if
	
	If locblnDisplayCancelRequestForm then
		mWriteCancelLeavePeriodRequestForm objEEtoView, locobjLeavePeriod, locblnValidateRequest
	End If
	
	If locblnDisplaySummary then
		mWriteViewingEmployee objEEtoView
		mWriteLeaveSummary objEEtoView, CONST_APPLICATION_PATH & "/leavesummary.asp"
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
