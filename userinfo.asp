<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
    strCurrentPageName = "User Info"
    '**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
    mInitialiseCurrentUser

    Dim shareWithTeamCalendar
    Dim wwid
    Dim locCmd
    Dim locRS  

    Set locCmd = Server.CreateObject("ADODB.Command")
    Set locCmd.ActiveConnection = glbConnection
    locCmd.CommandType = adCmdStoredProc

    If request.form("formname") = "frmShareWithTeamCalendar" then
        shareWithTeamCalendar = request.form("fldblnShare") 
        wwid = request.form("wwid")       

        locCmd.Parameters.Append locCmd.CreateParameter("strEEWWID", adWChar, adParamInput, 8, wwid)   
        if shareWithTeamCalendar = "on" then
            locCmd.CommandText = "usp_share_with_team_calendar_on"
        else
            locCmd.CommandText = "usp_share_with_team_calendar_off"
        end if

        Set locRS = locCmd.Execute   
    end if
    
    mWriteHMTLTop strCurrentPageName
    mWriteNavBar strCurrentPageName
    
    Select Case strMode
        Case "ph"
            mWritePublicHolidays
        Case Else
            mWriteUserInfo
    End Select
    mWritePageFooter
    
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
