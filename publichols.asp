<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!-- #include virtual="/eVacation/common/objects/calendar.asp" -->
<%
	strCurrentPageName = "Addition Page"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView
	
	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName
		
		'**GIVES THE USER THE OPTION TO ADD NEW HOLIDAYS
		mWriteAddHolidayForm
		
		'**GIVES THE USER THE OPTION TO DELETE A HOLIDAY ON A PARTICULAR DAY
		mWriteRemoveHolidayForm
		
		'**SHOWS ALL THE PUBLIC HOLIDAYS IN THE DATABASE AT THE MOMENT
		mWritePublicHolidays
	
	mWritePageFooter
%>
