<%  Option Explicit 
    
    Response.Expires = -1000
    Response.ExpiresAbsolute = Now() - 1
    Response.AddHeader "cache-control", "private"
    Response.AddHeader "pragma", "no-cache"
%>
<!--#include virtual="/eVacation/common/adovbs.asp" -->
<!--#include virtual="/eVacation/common/functions.asp" -->
<!--#include virtual="/eVacation/common/objects/objectmodel.asp" -->
<!--#include virtual="/eVacation/common/structure/structure.asp" -->
<%    
    
    '*** CLOSE APPLICATION ***
    function mCloseApplication
        Set glbPublicHolidays = nothing
        Set objEEtoView = nothing
        Set objCurrentUser = nothing

        glbConnection.close
        Set glbConnection = nothing
    end function

	'*** TESTING APPLICATION *** 
	CONST CONST_TEST_MODE = false	' false: - all development features listed below are turned off
	                                '  true: - allows CONST_DEVELOPER_ALIAS to change user at will (with dropdown)
    CONST CONST_UNDER_CONSTRUCTION = false ' true - diverts all users except user with CONST_DEVELOPER_ALIAS to the "Under Construction" page
	CONST CONSTDEBUGON = false 		' Prints all debug information (only to set developer)

    CONST RENEW_ALL_RECORDS_AND_CLEAN_DB = false 'renew all records in db and remove all terminated ones
	
	'************* Change developer alias to become admin
	
	CONST CONST_DEVELOPER_ALIAS = "GER\vkarpenk"
	CONST CONST_SEND_MAIL_TO_DEV = true     
	CONST CONST_DEVELOPER_EMAIL = "veronika.karpenko@intel.com"
	
    Public Sub mDebugPrint(ByVal locstrMessage)
        If CONSTDEBUGON then
            response.write locstrMessage
        End If
    End Sub
	
    '**** CONSTANTS: DATABASE ****
	CONST CONST_ADO_EVACATION_CONNECTION_STRING = "DRIVER=SQL Server;SERVER=,1433;DATABASE=;UID=;PWD="
    CONST CONST_WDS_CONNECTION_STRING = "DATA SOURCE=DenodoODBCdev;PROVIDER=MSDASQL"

    CONST CONST_FIRST_YEAR_SYSTEM_ACTIVE = 2006
    CONST CONST_APPLICATION_FEEDBACK_EMAIL = "veronika.karpenko@intel.com"'[MFILLAST 08-2006]
    CONST CONST_APPLICATION_FOOTER = "&copy; Intel Corporation"
	CONST CONST_APPLICATION_PATH = "/eVacation" ' : change path
	CONST CONST_MAIL_SERVER = "mailhost.ir.intel.com"
	
	'**** CONSTANTS: LEAVE ****   
	
    '***** CONST CONST_DATE_CONFIRMING_LEAVE_BEGINS = "January 1, 2009" 
    
    CONST CONST_DATE_CONFIRMING_LEAVE_BEGINS = 2009 
	
	CONST CONST_MAX_ANNUAL_COMP_DAYS = 3    '**********************COMP TIME**************************************
	CONST CONST_COMP_TIME_EXPIRY_DAYS = 1000 ' 12/01/2012 - rthyesx - Increased expiry from 30 to 1000 on info from drogers
	
	CONST CONST_MIN_ANNUAL_COMP_DAYS = 0

    '**** CONSTANTS: SECURITY ****
    CONST CONST_LOGON_SUCCESSFUL = 0
    CONST CONST_LOGON_ERROR_USER_BLANK = 1
    CONST CONST_LOGON_ERROR_USER_NOT_FOUND = 2
    CONST CONST_LOGON_ERROR_USER_ACCESS_DENIED = 3
    CONST CONST_LOGON_ERROR_USER_SET_UP_REQUIRED = 4
    CONST CONST_USER_PAGE_ACCESS_DENIED = 5
    CONST CONST_USER_NOT_ALLOWED_TO_DELEGATE = 6


    Dim CONST_USER_ERROR_TITLE(5)
    CONST_USER_ERROR_TITLE(CONST_LOGON_ERROR_USER_BLANK) = "User Logon Error"
    CONST_USER_ERROR_TITLE(CONST_LOGON_ERROR_USER_NOT_FOUND) = "User Logon Error"
    CONST_USER_ERROR_TITLE(CONST_LOGON_ERROR_USER_ACCESS_DENIED) = "User Logon Error"
    CONST_USER_ERROR_TITLE(CONST_LOGON_ERROR_USER_SET_UP_REQUIRED) = "User Logon Error"
    CONST_USER_ERROR_TITLE(CONST_USER_PAGE_ACCESS_DENIED) = "Access to Page Denied"


    '**** CONSTANTS: E-MAIL ****
    CONST CONST_EMAIL_SYSTEM_EMAIL_FROM = "eVacation_SIE@intel.com"'[MFILLAST 08-2006]
    CONST CONST_EMAIL_TYPE_EE_NOTIFICATION = 1
    CONST CONST_EMAIL_TYPE_APPOINTED_APPROVER = 2
    CONST CONST_EMAIL_TYPE_MANAGER_INFORMATION = 3
    CONST CONST_EMAIL_TYPE_REQUEST_RESPONSE = 4
    CONST CONST_EMAIL_TYPE_EE_CANCELLATION_REQUEST_NOTIFICATION = 5
    CONST CONST_EMAIL_TYPE_CANCELLATION_REQUEST_APPOINTED_APPROVER = 6
    CONST CONST_EMAIL_TYPE_CANCELLATION_REQUEST_MANAGER_INFORMATION = 7
	CONST CONST_EMAIL_TYPE_EE_CONFIRM_LEAVE_TAKEN = 8 ' [MOFLYNN 11-2008] Added new email
    CONST CONST_EMAIL_LEAVE_BACKGROUND = "#FFE4C4"
    

    '**** CONSTANTS: FORM ERRORS ****
    CONST CONST_FORM_ERROR_NOT_COMPLETED_CORRECTLY = 1
    CONST CONST_FORM_ERROR_INVALID = 2

    
    '**** CONSTANTS: OBJECT MODEL - LEAVE TYPES ****
    CONST CONST_LEAVE_TYPE_NAME_ANNUAL_VACATION = "Annual Leave"
    CONST CONST_LEAVE_TYPE_NAME_ELP = "ELP Vacation"
    CONST CONST_LEAVE_TYPE_NAME_COMP_TIME = "Comp Leave"
    

    '**** CONSTANTS: OBJECT MODEL - LEGAL ADJUSTMENT TYPES ****
    CONST CONST_LEGAL_ADJUSTMENT_TYPE_NAME_REST_DAYS = "Rest Days"
    CONST CONST_LEGAL_ADJUSTMENT_TYPE_NAME_JRTT = "JRTT"


    '**** CONSTANTS: OBJECT MODEL - LEGAL ADJUSTMENT SETTINGS ****
    CONST CONST_LEGAL_ADJUSTMENT_MAX_WORKING_DAYS = 218

    
    '**** CONSTANTS: OBJECT MODEL - ELP SETTINGS ****
    CONST CONST_ELP_DAYS_BANKED_PER_YEAR = 5
    CONST CONST_ELP_TARGET_DAYS = 30
    CONST CONST_ELP_STATUS_ACTIVE = "Active"
    CONST CONST_ELP_STATUS_MATURED = "Matured"
    CONST CONST_ELP_STATUS_USED = "Used"
    CONST CONST_ELP_STATUS_EXPIRED = "Expired"
    CONST CONST_ELP_RELIEF_ACTION_ADD = 1
    CONST CONST_ELP_RELIEF_ACTION_REMOVE = 0

    '**** CONSTANTS: OBJECT MODEL - EMPLOYEE COLLECTION TYPES ****
    CONST CONST_EMPLOYEE_COLLECTION_TYPE_DIRECT_REPORTS = 1
    CONST CONST_EMPLOYEE_COLLECTION_TYPE_NAME_SEARCH = 2
    CONST CONST_EMPLOYEE_COLLECTION_TYPE_PAYROLL_REPORT = 3


    '**** CONSTANTS: OBJECT MODEL - LEAVE TYPE COLLECTION TYPES ****
    CONST CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_EE = 1
    CONST CONST_LEAVE_TYPE_COLLECTION_TYPE_LEAVE_REQUESTS_FOR_ADMIN = 2
    CONST CONST_LEAVE_TYPE_COLLECTION_TYPE_OTHER_LEAVE_TYPES_FOR_EE = 3
    

    '**** CONSTANTS: OBJECT MODEL - LEAVE PERIOD COLLECTION TYPES ****
    CONST CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_REQUESTS = 1
    CONST CONST_LEAVE_PERIOD_COLLECTION_TYPE_LEAVE_APPROVALS = 2
    CONST CONST_LEAVE_PERIOD_COLLECTION_TYPE_ANNUAL_LEAVE = 3
    CONST CONST_LEAVE_PERIOD_COLLECTION_TYPE_ALL_EE_LEAVE_IN_PERIOD = 4
    CONST CONST_LEAVE_PERIOD_COLLECTION_TYPE_OTHER_LEAVE = 5
    CONST CONST_LEAVE_PERIOD_COLLECTION_TYPE_ADMIN_VIEW = 6


    '**** CONSTANTS: OBJECT MODEL - LEAVE PERIOD STATUS ****
    CONST CONST_LEAVE_PERIOD_STATUS_RAISED = "Pending Approval"
    CONST CONST_LEAVE_PERIOD_STATUS_APPROVED = "Leave Approved"
    CONST CONST_LEAVE_PERIOD_STATUS_REJECTED = "Rejected"
    CONST CONST_LEAVE_PERIOD_STATUS_CANCEL_REQUESTED = "Cancel Requested"
    CONST CONST_LEAVE_PERIOD_STATUS_CANCEL_APPROVED = "Cancelled"
    CONST CONST_LEAVE_PERIOD_STATUS_CANCEL_REJECTED = "Cancel Rejected"
	CONST CONST_LEAVE_PERIOD_STATUS_CONFIRMED = "Confirmed as Taken"    


    '**** CONSTANTS: OBJECT MODEL - COMP. TIME STATUS ****  
    CONST CONST_COMP_TIME_STATUS_GRANTED = "Granted"
    CONST CONST_COMP_TIME_STATUS_REVOKED = "Revoked"
    CONST CONST_COMP_TIME_STATUS_EXPIRED = "Expired"
    CONST CONST_COMP_TIME_STATUS_BOOKED  = "Booked"
    CONST CONST_COMP_TIME_STATUS_USED    = "Taken"

    '**** CONSTANTS: APPLICATION MODES *****
    CONST CONST_MODE_APPROVE_REQUESTS_VIEW_REQUEST = "approverequest"


    '**** VARIABLE DECLARATIONS ****
    Public glbConnection
    Public objCurrentUser
    Public objEEtoView 
    Public glbPublicHolidays ' list of public holidays
    Dim lngLoadCurrentUserStatus
    Dim strRequestEEWWID
    Dim strMode 
    Dim strCurrentPageName
    Dim lngItemID
    
    Dim glbObjectCounter
    Dim glbObjectTerminateCounter
    
    glbObjectCounter = 0
    glbObjectTerminateCounter = 0

    '**** GLOBAL CONNECTION VARIABLE ****

    Set glbConnection = Server.CreateObject("ADODB.Connection")
    
    glbConnection.ConnectionString = CONST_ADO_EVACATION_CONNECTION_STRING
    glbConnection.Open


    '**** PUBLIC HOLIDAYS COLLECTION ****
    Set glbPublicHolidays = new cColPublicHolidays
    
    '**** RETRIEVE REQUEST VARIABLES ****

    strRequestEEWWID = request("ee")
    strMode = request("m")
    lngItemID = mGetSafeLongInteger(request("itemid"),0)

%>
