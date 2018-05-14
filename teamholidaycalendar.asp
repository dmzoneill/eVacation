<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!--#include virtual="/eVacation/common/objects/calendar.asp" -->
<%
	strCurrentPageName = "Employee Leave Calendar"

    Dim locCmd
    Dim locRS 

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView

	If not objCurrentUser.IsEELeaveTracked then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If

	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

    if lngItemID then

        Set locCmd = Server.CreateObject("ADODB.Command")
        Set locCmd.ActiveConnection = glbConnection
        locCmd.CommandType = adCmdStoredProc

        locCmd.Parameters.Append locCmd.CreateParameter("lngID", adWChar, adParamInput, 8, lngItemID) 
            If strMode = "htcl" then
                locCmd.CommandText = "usp_leave_share_with_team_calendar_off"
            elseif  strMode = "stcl" then
                 locCmd.CommandText = "usp_leave_share_with_team_calendar_on"
            end if

        Set locRS = locCmd.Execute   
    end if

	mWriteCalendarHolidayDisplay objEEtoView, True

	mWritePageFooter
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
