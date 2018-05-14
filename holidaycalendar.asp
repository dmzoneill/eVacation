<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!-- #include virtual="/eVacation/common/objects/calendar.asp" -->
<%
	strCurrentPageName = "Employee Leave Calendar"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView

	If not objCurrentUser.IsEELeaveTracked then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_PAGE_ACCESS_DENIED
	End If

	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

	mWriteCalendarHolidayDisplay objEEtoView, False

	mWritePageFooter
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
