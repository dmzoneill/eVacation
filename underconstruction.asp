<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	mWriteHMTLTop "e-Vacation - Under Construction"
	

		response.write "<table width=700 border=0 cellspacing=0 cellpadding=3 align=center>"
			response.write "<tr>"
				response.write "<td align=""center"">"
					response.write "<img style=""width:200px;"" src=""" & CONST_APPLICATION_PATH & "/common/images/uc.jpg"" alt=''>"
				response.write "</td>"
			response.Write "</tr>"	
	    response.Write "</table>"
	
	    response.Write "<br class=small>"
	    
		response.write "<table width=700 border=0 cellspacing=0 cellpadding=3 align=center class=tbcontent>"
			response.Write "<tr>"
				response.write "<td>"
					response.write "<br class=small>"
				response.write "</td>"
				response.write "<td class=txttitle style=""text-align: center"">"
					response.write "e-Vacation is Under Construction"
				response.write "</td>"
				response.write "<td>"
					response.write "<br class=small>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
	
	    response.Write "<br class=small>"
	
		response.write "<table width=700 border=0 cellspacing=0 cellpadding=3 align=center class=tbcontent>"
			response.Write "<tr>"
				response.write "<td>"
					response.write "<br class=small>"
				response.write "</td>"
				response.write "<td style=""text-align: center"">"
					response.write "<br><br>"
					response.write "The tool will be operational again at 1PM today. Apologies for any inconvenience caused."
					response.write "<br><br>"
					response.write "For any urgent issues please contact the developer: " & CONST_DEVELOPER_EMAIL
					response.write "<br><br>"
				response.write "</td>"
				response.write "<td>"
					response.write "<br class=small>"
				response.write "</td>"
			response.write "</tr>"
		response.write "</table>"
	
	    response.Write "<br class=small>"
	
	mWritePageFooter
	
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
