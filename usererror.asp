<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	strCurrentPageName = "Error"

	Dim loclngError

	loclngError = mGetSafeLongInteger(request.querystring("error"),0)
	
	mWriteHMTLTop strCurrentPageName
	
	mWriteUserError(loclngError)
	mWritePageFooter
	
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
