<%
'===============================================================
'	sendconfirmemail.vbs
'
'	Author: 	 Matthew O'Flynn
'	Created:	 4/11/2008
'	Description: This script searches for employees who had holidays
'				 that ended yesterday. Each of these employees are 
'				 e-mailed asking them to confirm that they did 
'				 take the holidays. 
'===============================================================
%>
<!-- #include virtual="eVacation/common/appglobal.asp" -->
<%

'*** SCRIPT DETAILS ***
Dim StartTime
StartTime = Timer()
Server.ScriptTimeout = 300 ' 5 minutes
response.Write "<b>***************<br>"
response.write "Script Details:</b><br>"
response.write "Script Timeout: " & Server.ScriptTimeout & "<br><br>"
response.Write ""


'*** FUNCTION LIST ***
' mGetAllActiveEmployees	
' mGetEELeaveRequestsToConfirm(WWID)

'*** GET ALL ACTIVE EMPLOYEES ***
'Returns collection of WWIDs of all employees that are leave tracked from DB
function mGetAllActiveEmployees
	Dim locarrReturnValue() ' Dynamic array for WWID number of employees
	Dim sqlGetAllActiveEEs 
	Dim locGetAllActiveEECommand
	Dim locRS ' RecordSet
	Dim locCount 

	' Print Debug Info
	response.write "<b>*******************************************<br>"
	response.write " LOADING ACTIVE EMPLOYEES<br></b>"
	
	' Load WWIDs from DB 
	sqlGetAllActiveEEs = "SELECT WWID " & _
						 "FROM WorkerPrivate wp, tblUser u " & _
						 "WHERE wp.WWID = u.strWWID and u.blnIsEELeaveTracked<>0 and wp.EmployeeStatusCd='A'"
'	Set locGetAllActiveEECommand = Server.CreateObject("ADODB.Command")
'	locGetAllActiveEECommand.ActiveConnection = glbConnection
'	locGetAllActiveEECommand.CommandText = sqlGetAllActiveEEs
	
'	Set locRS = locGetAllActiveEECommand.Execute
	
	Set locRS = Server.CreateObject("ADODB.RecordSet")
	locRS.Open sqlGetAllActiveEEs, _
			   glbConnection, _
			   adOpenStatic
	
	
	' Check result
	If locRS.EOF Then
		response.write "<b>Failed to get any names...<br></b>"
		reDim locarrReturnValue(1)
		locarrReturnValue(0) = 0
	Else
		response.write "<b>Loading names... [" & locRS.RecordCount & " Names]</b>"

		' Create an array with just enough space
		reDim locarrReturnValue(locRS.RecordCount)
		
		' Save all the WWIDs
		locCount = 0
		While not locRS.EOF	
			locarrReturnValue(locCount) = locRS("WWID")
			
			locCount = locCount + 1
			locRS.MoveNext
		Wend
		
		response.write "<b>done<br></b>"
	End If

    ' Clean up
    locRS.Close
    Set locRS = nothing

	mGetAllActiveEmployees = locarrReturnValue
end function 


'*** GET EMPLOYEE LEAVE REQUESTS ***
' Returns a collection of Leave Requests for given employee 
' which the employee can confirm
function mGetEELeaveRequestsToConfirm(WWID)
	' Set up current user
	' We expect that ee (or WWID) will be set in the GET request,
	' so we can use this to initialise the current user
	Dim objUser
	Dim objUserPreviousYear
	Dim colLeave  ' collection of leave requests
	Dim colApprovedLeave
	Dim objLeavePeriod
	Dim loclngCount
	Dim locdatConfirmLeaveFeatureBegin ' Date after which Leave Requests need to be confirmed

	locdatConfirmLeaveFeatureBegin = cDate(CONST_DATE_CONFIRMING_LEAVE_BEGINS)
	
	Set objUser = new cObjUser
	Set objUserPreviousYear = new cObjUser
	objUser.WWID=WWID ' Set the WWID so that we can load the user
	objUserPreviousYear.WWID = WWID
	objUser.YearToView = DatePart("yyyy", Now()) ' Use current year
	objUserPreviousYear.YearToView = objUser.YearToView - 1 ' Use last year also
	
	response.Write "<td>" & WWID & "</td>"
	response.write "<td>" & objUser.FirstNm & " " & objUser.LastNm & "</td>"
	
	' Create collection to store only approved leave requests
	Set colApprovedLeave = new cObjCollection
	
	'**** Get This Years Leave Requests
	' Load user data (leave periods is all we want)
	Set colLeave = objUser.LeaveRequests
	
	response.write "<td>" ' Leave requests cell
	
	' Check if there is no leave requests
	If colLeave.Count <> 0 Then
	    ' For each Leave Request
	    For loclngCount = 1 To colLeave.Count
		    Set objLeavePeriod = colLeave.Item(loclngCount)
    	
    	    ' Check if leave periods are "approved", have ended and the email has not been sent yet
    	    ' If an employee has submitted a cancellation request and it has been rejected by the manager another "confirm" mail is sent to employee
		    If (objLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_APPROVED or _
		        objLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED) and _
		       objLeavePeriod.CanConfirmLeave and _
		       objLeavePeriod.EndDate > locdatConfirmLeaveFeatureBegin and _
		       not isDate(objLeavePeriod.DateConfirmEmailSent) Then
	    		colApprovedLeave.Add objLeavePeriod
			    response.write "<span style='color: green;'>[" & objLeavePeriod.StartDate & "-" & objLeavePeriod.EndDate & "]</span> "
		    'Else
			    'response.write "<span style='color: red;'>[" & objLeavePeriod.EndDate & ": " & objLeavePeriod.Status & "]</span> "
		    End If
	    Next
	End If 
	
	'**** Include the Leave Requests for Last Year also
	Set objLeavePeriod = nothing
	Set colLeave = objUserPreviousYear.LeaveRequests
	
	If colLeave.Count <> 0 Then
		' For each Leave Request
	    For loclngCount = 1 To colLeave.Count
		    Set objLeavePeriod = colLeave.Item(loclngCount)

    	    ' Check if leave periods are "approved", have ended and the email has not been sent yet
		    If objLeavePeriod.Status = CONST_LEAVE_PERIOD_STATUS_APPROVED and _
		       objLeavePeriod.CanConfirmLeave and _
		       objLeavePeriod.EndDate > locdatConfirmLeaveFeatureBegin and _
		       not isDate(objLeavePeriod.DateConfirmEmailSent) Then
	    		colApprovedLeave.Add objLeavePeriod
			    response.write "<span style='color: green;'>[" & objLeavePeriod.EndDate & ": " & objLeavePeriod.Status & "]</span> "
'		    Else
'			    response.write "<span style='color: red;'>[" & objLeavePeriod.EndDate & ": " & objLeavePeriod.Status & "]</span> "
		    End If
	    Next
	End If

	response.write "</td>" ' Leave requests cell	
	response.Write "<td>" & colApprovedLeave.Count & "</td>"
	
	' Return array of leave requests
	Set mGetEELeaveRequestsToConfirm = colApprovedLeave
End Function


'*** LOGIC ***
' Send out emails to each employee
Dim WWIDs           '
Dim Count           '
Dim lrCount         '
Dim LeaveRequests   '
Dim LR              '
Dim blnMailSent
Dim sqlMailSent
WWIDs = mGetAllActiveEmployees

Dim emailCount
emailCount = 0

' Print table headers
response.write "<table>"
	response.write "<tr style=""text-align: left;"">"
		response.write "<th>WWID</th>"
		response.write "<th>Name</th>"
		response.write "<th>Leave Requests to Confirm:</th>"
		response.write "<th>Total E-mails</th>"
		response.write "<th>Emails Status</th>"
	response.write "</tr>"

' For each WWID (Employee)
For Count = 0 to uBound(WWIDs)-1
'	If WWIDs(Count) = 11229520 Then     ' Debug condition
		response.write "<tr>"
	
	    Set LeaveRequests = mGetEELeaveRequestsToConfirm(WWIDs(Count))

		response.write "<td>"   	
	    For lrCount = 1 to LeaveRequests.Count 
		    emailCount = emailCount + 1 ' count emails sent		
		    Set LR = LeaveRequests.Item(lrCount)

		    ' SEND AN EMAIL FOR THIS LEAVE REQUEST
	'	    response.write "Sending email to " & LR.EE.Email & " from eVacation@intel.com<br>"
	'	    response.Write "Subject: ""e-Vacation - " & LR.EE.FirstNm & " " & LR.EE.LastNm & " - Leave Request to Confirm""<br>"
	'	    response.write "Email:<br>" & LR.LeaveRequestEmailBody(CONST_EMAIL_TYPE_EE_CONFIRM_LEAVE_TAKEN)
		   
		    ' Use existing mSendMail function to send mail to employee
		    blnMailSent = mSendEmail("evacation_sie@intel.com", _ 
		                             LR.EE.Email, _ 
		                             "e-Vacation - Leave Request to Confirm", _ 
		                             LR.LeaveRequestEmailBody(CONST_EMAIL_TYPE_EE_CONFIRM_LEAVE_TAKEN), _ 
		                             true)
           
			' Check email was sent... and mark it as sent in DB
			If not blnMailSent Then
			    response.Write "<b>Mail not sent!</b>"
			Else
			    response.Write "Mail sent! "
			    
			    ' Update the Leave Request with the date that the email was sent
			    sqlMailSent = "UPDATE tblLeavePeriod " & _
			                  "SET datConfirmEmailSent='" & Now & "' " & _
			                  "WHERE lngID=" & LR.ID
			    glbConnection.Execute sqlMailSent
			End If
	    Next
		response.write "</td>"
		
		response.write "</tr>"		
'    End If
Next

response.write "</table>"

response.write "Sent " & emailCount & " emails in " & Timer()-StartTime & " seconds."
%>