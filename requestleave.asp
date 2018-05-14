<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!-- #include virtual="/eVacation/common/objects/calendar.asp" -->
<%
	strCurrentPageName = "Request Leave"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView

	If not objCurrentUser.IsEELeaveTracked then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If
	
	Dim locobjLeaveRequest
	Dim locblnValidateRequest
	Dim loclngSaveResult
    Dim locblnShowRequestLeaveForm
    Dim locblnShowCompTimeDetails
    Dim locblnDaysRequested
    Dim locobjCompTime
    Dim locblnEnoughDaysInCompTime
	
	locblnShowRequestLeaveForm = True
	locblnDaysRequested = False
	locblnEnoughDaysInCompTime = False
	
	Set locobjLeaveRequest = new cObjLeavePeriod
	
	If request.form("formname") = "frmRequestLeave" then
		locblnValidateRequest = True
		locobjLeaveRequest.LoadNewRequestFromForm
	Else
		locblnValidateRequest = False
	End If
	
	' Load Comp Time Details if in Comp Time mode
	If strMode = "ct" then
	    Set locobjCompTime = new cObjCompTime
        locobjCompTime.ID = lngItemID
        locblnShowCompTimeDetails = true
        
        ' Check if the days available in the comp time => the number of days requested
        if request.form("formname") = "frmRequestLeave" then
            if locobjCompTime.DaysAvailable >= locobjLeaveRequest.Days then
                locblnEnoughDaysInCompTime = True
            end if
        end if        
	End If 
	
	

	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

	if locblnValidateRequest then
		if locobjLeaveRequest.FormIsValid then
			if locobjLeaveRequest.NewRequestIsValid then
			    ' Save Comp Time ID if request is for Comp Time
			    if locobjLeaveRequest.LeaveType.Name = CONST_LEAVE_TYPE_NAME_COMP_TIME then
			        locobjLeaveRequest.CompTimeID = Request.Form("itemid")
			    end if 
			    			    
			    ' Check if Days are available if this is for Comp Time
			    if not strMode = "ct" or locblnEnoughDaysInCompTime then
				    loclngSaveResult = locobjLeaveRequest.Save
    				locblnDaysRequested = true
    				
				    If loclngSaveResult = 0 then
					    mWriteGeneralError "Sorry - an error has occurred while attempting to save your leave request.<br>" & _
						    "There may be a problem with the network, or the database server may be experiencing difficulties.", False
				    else
					    response.write "<div class='feedbackbox'>"
						    response.write "<b>Your leave request has been raised successfully.</b>"
					    response.write "</div>"
					    Set locobjLeaveRequest = nothing			
					    Set locobjLeaveRequest = new cObjLeavePeriod
					    locblnValidateRequest = False
					    
					    ' Check if there are any Days left in this Comp Time
					    If strMode = "ct" then
					        If locobjCompTime.DaysAvailable = 0 then
					            locblnShowRequestLeaveForm = False
					        End If
					    End If
				    end if
				end if
			end if
		end if
	end if
	
	' If we are booking comp time
	if strMode = "ct" then
        if not locobjCompTime.IsLoaded then 
            mWriteGeneralError "Sorry - there is no comp time with that ID", False
            locblnShowCompTimeDetails = false
            locblnShowRequestLeaveForm = false
        else
            if locobjCompTime.DaysAvailable = 0 and not locblnDaysRequested then
                mWriteGeneralError "Sorry - there are no days available for the comp time specified.", False
                locblnShowRequestLeaveForm = false
            end if
        end if

	    locobjLeaveRequest.CompTimeID = lngItemID
	end if
  
    ' Print Comp Time Details
    if locblnShowCompTimeDetails then
        mWriteCompTimeInstanceDetails lngItemID
    end if
    
    ' Print Request Leave Form
    if locblnShowRequestLeaveForm then
        mWriteLeaveRequestForm objEEtoView, locobjLeaveRequest, locblnValidateRequest
    end if
    
    Set locobjCompTime = nothing
	Set locobjLeaveRequest = nothing

	mWritePageFooter
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
