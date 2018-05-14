<!--#include virtual="/eVacation/common/appglobal.asp" -->
<%
	strCurrentPageName = "Delegate"

	'**** INITIALISE CURRENT USER AND EE TO VIEW OBJECTS ****
	mInitialiseCurrentUser
	
	Dim loclngSaveResult
	Dim locblnValidate

	If not objCurrentUser.IsManager then
		mCloseApplication
		response.redirect CONST_APPLICATION_PATH & "/usererror.asp?error=" & CONST_USER_NOT_ALLOWED_TO_DELEGATE
	End If

	mWriteHMTLTop strCurrentPageName
	mWriteNavBar strCurrentPageName

	locblnValidate = False

	If request.form("formname") = "frmRemoveDelegate" then
		loclngSaveResult = objCurrentUser.RemoveDelegate()
		if loclngSaveResult = 0 then
			mWriteGeneralError "Sorry - an error has occurred while attempting to remove your delegate.<br>" & _
				"There may be a problem with the network, or the database server may be experiencing difficulties.", False
		else
			response.write "<center>"
				response.write "<br>"
				response.write "<b>Your delegate has been removed successfully.</b><br>"
				response.write "<br>"
			response.write "</center>"
		End If
	ElseIf request.form("formname") = "frmAppointDelegate" then
		objCurrentUser.LoadDelegateFromForm
		If objCurrentUser.DelegateFormIsValid then
			loclngSaveResult = objCurrentUser.Save()
			If loclngSaveResult = 0 then
				mWriteGeneralError "Sorry - an error has occurred while attempting to appoint your delegate.<br>" & _
					"There may be a problem with the network, or the database server may be experiencing difficulties.", False
			else
				response.write "<center>"
					response.write "<br>"
					response.write "<b>You have successfully appointed "
					response.write objCurrentUser.ActiveDelegate.FullName
					response.write " as your delegate.</b><br>"
					response.write "<br>"
				response.write "</center>"
			End If
		Else
			locblnValidate = True
		End If
	End If

	mWriteDelegateForm objCurrentUser, locblnValidate
	
	mWritePageFooter
%>
<!--#include virtual="/eVacation/common/appglobalend.asp" -->
