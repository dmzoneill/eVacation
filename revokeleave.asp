<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!-- #include virtual="/eVacation/common/objects/calendar.asp" -->
<%
	strCurrentPageName = "Revoke Leave"  '[CH 02/09]

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView
	
	' TO DO: Check that this is the manager of the employee (permissions)

	If not objCurrentUser.IsEELeaveTracked then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If
	
	Dim locblnMaxCompTimeGranted
	Dim locblnMinCompTimeGranted
	Dim locblnValidateRequest
	Dim loclngSaveResult
    Dim locObjCompTime
    Set locObjCompTime = new cObjCompTime
	
	If request.form("formname") = "frmRequestLeave" then
		locblnValidateRequest = True
        locObjCompTime.LoadFromRevokeForm
        locObjCompTime.Revoke
	Else
		locblnValidateRequest = False
	End If

	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

    ' Check if the Max Comp Time has already been granted
    locObjCompTime.LoadEE
    If objCurrentUser.WWID = locObjCompTime.EE.Manager.WWID then
        If objEEtoView.AnnualVacation.TotalCompTimeGranted = 0 then
	        locblnMinCompTimeGranted = true
        
		    response.write "<br><center>"
                response.write "<font class=error>"
                    response.write "<b>Sorry - this employee (" & objEEtoView.WWID & ") has no Comp Days to revoke (" & CONST_MIN_ANNUAL_COMP_DAYS & " Days).<br>"
                    response.write "<br>"
                    response.write "Comp Days can be granted to this employee again once the current Comp Days are Taken, Revoked or Expire.</b><br>"
                    response.write "<br><br>"
                response.write "</font>"
            response.write "</center>"
            
        Else 
            locblnMinCompTimeGranted = false
            mWriteRevokeLeaveForm objEEtoView, locblnMinCompTimeGranted, locblnValidateRequest
        End If
    else
        mWriteGeneralError "Sorry - you do not have the authority to revoke Compensatory Time of this employee.", False 
    end if

'	Set locobjLeaveRequest = nothing

	mWritePageFooter
%>

<!--#include virtual="/eVacation/common/appglobalend.asp" -->
