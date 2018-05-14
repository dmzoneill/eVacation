<!--#include virtual="/eVacation/common/appglobal.asp" -->

<% 'start ASP code

'execute the query given in parameter("query") and display the result in a opup window
'developped by [MFILLAST 08-2006]




dim locstrQuery 'sql query to execute
'connection to database
Dim locCmd
Dim locParam
Dim locRS
Dim fld

	mInitialiseCurrentUser
	'**** CHECK THAT THE CURRENT USER HAS ADMINISTRATOR RIGHTS.
	If not objCurrentUser.IsAdmin then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If
	
	
locstrQuery = (Request.querystring("query")) 'TODO reencode query in original format,+ is deleted




' display the header with css stylesheet
%>
<html>
	<head>
 		<title>e-Vacation - MSSQL database query</title>
		<meta http-equiv=Content-Type content=text/html; charset=iso-8859-1>
		<meta name='description' content='Updating MSSQL database'>
		<meta name='author' content='Mikaël Fillastre'>
		<STYLE TYPE='text/css' >
		<!--
		.visibleTable, .visibleTable TR, .visibleTable TD, .visibleTable TH,
		{
		text-align: center;
		border-color: black;
		border-collapse: collapse;
		empty-cells:show;
		border-width: 1px;
		border-style: solid;
		font-size:10px;
		}
		.visibleTable TH {background-color:#cccccc}
		-->
		</STYLE>
	</head>

<table width=100% border=0 cellspacing=0 cellpadding=0>
	<tr align=center valign=top>
		<td height="300">
			<table width=550 border=0 cellspacing=0 cellpadding=0>
				<tr>
  					<td>
  						<br>


    					<table width=550 border=0 cellspacing=1 cellpadding=1 bgcolor=#aaaaff>
							<tr>
								<td>
									<font face="Verdana, Arial, Helvetica, sans-serif" size=2 color=#FFFFFF>
										<b>QUERY :</b>
										</font>
								</td>
      						</tr>
							<tr>
								<td height="43">
									<table width="546" border=0 cellspacing=1 cellpadding=1 bgcolor=#FFFFFF>
										<tr align="left" valign=top>
											<td>
												<font face="Verdana, Arial, Helvetica, sans-serif" size=1>
<%
response.Write locstrQuery
%>
<br>
												</font>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
						<br>

            
						
						   <table width=550 border=0 cellspacing=1 cellpadding=1 bgcolor=#aaaaff ID="Table1">
							<tr>
								<td>
									<font face="Verdana, Arial, Helvetica, sans-serif" size=2 color=#FFFFFF>
										<b>RESULT :</b>
										</font>
								</td>
      						</tr>
							<tr>
								<td height="43">
									<table width="546" border=0 cellspacing=1 cellpadding=1 bgcolor=#FFFFFF ID="Table2">
										<tr align="left" valign=top>
											<td>
												<font face="Verdana, Arial, Helvetica, sans-serif" size=1>
<%


'create the command
set locCmd = Server.CreateObject("ADODB.Command")
Set locCmd.ActiveConnection = glbConnection
locCmd.CommandText = locstrQuery

set locRS = locCmd.Execute

if (locRS.State=1) then '1=open
   if (locRS.EOF) then 
		response.write "no record correspond to the query"
   else
		' display the result in a table
		response.write "		<table class='visibleTable' >"
			response.write "		<tr>"
			For Each fld In locRS.Fields
			response.write "<th>"&fld.Name&"</th>"
			Next           
			response.write "		</tr>"
		While not locRS.eof
			response.write "		<tr>"
			For Each fld In locRS.Fields
				response.write "<td>"&fld.Value&"</td>"
			Next           
			response.write "		</tr>"
			locRS.movenext
		Wend
		response.write "		</table>"  
    End if
    ' close the connection
	locRS.Close
else 'nothing to display : insert, update...
	response.write "nothing to display"
end if


			

Set locRS = nothing
Set locParam = nothing
Set locCmd = nothing     
            

  ' end of html page  
%>

<br>
												</font>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
						<br>
						</body>
</html>
