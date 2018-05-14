<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	strCurrentPageName = "Team Leave Summary"
	
	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser	
	
	'dim endtime
	'endtime = timer()
	'dim benchmark
	'benchmark = endtime - starttime
	'Response.Write( benchmark ) 
	'Response.Write( " - teamleavesummary.asp" ) 
	'Response.Write( "<br>" )
	
	mInitialiseEEtoView	
	
	If Request.Form("update") = "yes" Then
		updateEmployees objEEtoView
    End if
	
	mWriteHMTLTop strCurrentPageName	
	mWriteNavBar strCurrentPageName
	
	Dim compTime
	
	'**** GRANT COMPENSATORY TIME TO EE ****
	If request.Form("formname") = "frmGrantCompTime" Then
	    
	    ' Load details from form (Grant Comp. Time Form)
	    Set compTime = new cObjCompTime
	    compTime.LoadFromGrantForm
	    
	    ' TO DO: Validate the form
	    ' TO DO: Check if user's total comp days will not go over CONST_MAX_ANNUAL_COMP_DAYS limit if this request is processed
	    	    
	    'response.Write "Status = " & compTime.Status & "<br>"
	
	    ' TO DO: Send an email to the employee notifying them about the comp. time
		
		
		' TO DO: Add a new "Comp. Time" leave request to the DB 
		compTime.Save
		
    	response.write "<center>"
			response.write "<br>"
			response.write "<b>You have successfully granted " & CompTime.EE.FirstNm & compTime.EE.LastNm & " (" & strRequestEEWWID & ") " & request("fldlngCompDays") & " day"
			If 1 <> request("fldlngCompDays") then
			    response.Write "s"
			End If
			response.write "Compensatory Leave</b><br>"
			response.write "<br>"
		response.write "</center>"
		
	End If
	
	'**** GRANT COMPENSATORY TIME TO EE ****
	If request.Form("formname") = "frmRevokeCompTime" Then
	    
	    ' Load details from form (Grant Comp. Time Form)
	    Set compTime = new cObjCompTime 
	    compTime.LoadFromGrantForm
	    
	    ' TO DO: Validate the form
	    ' TO DO: Check if user's total comp days will not go over CONST_MAX_ANNUAL_COMP_DAYS limit if this request is processed
	    	    
	    'response.Write "Status = " & compTime.Status & "<br>"
	
	    ' TO DO: Send an email to the employee notifying them about the comp. time
		
		
		' TO DO: Add a new "Comp. Time" leave request to the DB 
		compTime.Revoke
		
    	response.write "<center>"
			response.write "<br>"
			response.write "<b>You have successfully revoked " & CompTime.EE.FirstNm & compTime.EE.LastNm & " (" & strRequestEEWWID & ") " & request("fldlngCompDays") & " day"
			If 1 <> request("fldlngCompDays") then
			    response.Write "s"
			End If
			response.write " Compensatory Leave</b><br>"
			response.write "<br>"
		response.write "</center>"
		
	End If


	mWriteTeamLeaveSummary objEEtoView

	mWritePageFooter
%>

<!--#include virtual="/eVacation/common/appglobalend.asp" -->
