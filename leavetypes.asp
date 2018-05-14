<!--#include virtual="/eVacation/common/appglobal.asp" -->
<!-- #include virtual="/eVacation/common/objects/calendar.asp" -->
<%
	strCurrentPageName = "Addition Page"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	mInitialiseEEtoView
	
	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName
		
		'**FORM FOR ADDING NEW LEAVE TYPES
		mWriteAddLeaveType
		
		'**FORM FOR REMOVING LEAVE TYPES
		mWriteRemoveLeaveType
	
	mWritePageFooter
%>
