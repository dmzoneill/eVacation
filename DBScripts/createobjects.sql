


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblELPRelief_tblELPInstance]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblELPRelief] DROP CONSTRAINT FK_tblELPRelief_tblELPInstance
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblLeavePeriod_tblLeaveType]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblLeavePeriod] DROP CONSTRAINT FK_tblLeavePeriod_tblLeaveType
GO



/* proc created by [MFILLAST 08-2006] */
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_insert_cdis_copy]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_insert_cdis_copy]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_update_cdis_copy]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_update_cdis_copy]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_getLastUpdateCdisCopy_from_wwid]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_getLastUpdateCdisCopy_from_wwid]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_getLastUpdateCdisCopy_from_shortid]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_getLastUpdateCdisCopy_from_shortid]
GO

/* end of proc created by [MFILLAST 08-2006] */

/****** Object:  Stored Procedure dbo.pr_evc_elp_data_for_employee    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_elp_data_for_employee]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_elp_data_for_employee]
GO

/****** Object:  Stored Procedure dbo.pr_evc_employee_carry_over_for_year    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_employee_carry_over_for_year]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_employee_carry_over_for_year]
GO

/****** Object:  Stored Procedure dbo.pr_evc_employee_leave_for_year    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_employee_leave_for_year]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_employee_leave_for_year]
GO

/****** Object:  Stored Procedure dbo.pr_evc_shannon_sites    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_shannon_sites]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_shannon_sites]
GO

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_base_data    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_hr_report_base_data]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_hr_report_base_data]
GO

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_employee_data    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_hr_report_employee_data]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_hr_report_employee_data]
GO

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_manager_reports    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_hr_report_manager_reports]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_hr_report_manager_reports]
GO

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_site_only    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_payroll_report_filtered_site_only]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_payroll_report_filtered_site_only]
GO

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_with_site    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_payroll_report_filtered_with_site]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_payroll_report_filtered_with_site]
GO

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_without_site    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[pr_evc_payroll_report_filtered_without_site]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[pr_evc_payroll_report_filtered_without_site]
GO

/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_approve_cancel_request]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_approve_cancel_request]
GO

/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_approve_leave_request]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_approve_leave_request]
GO

/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_cancel_leave_request]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_cancel_leave_request]
GO

/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_carryover_by_ID]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_carryover_by_ID]
GO

/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_carryover_by_WWID_Year]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_carryover_by_WWID_Year]
GO

/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_carryovers_by_WWID]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_carryovers_by_WWID]
GO

/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_create_user]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_create_user]
GO

/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_delete_leave_request]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_delete_leave_request]
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ee_payrollreport]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ee_payrollreport]
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport_old    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ee_payrollreport_old]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ee_payrollreport_old]
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ee_payrollreportlite]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ee_payrollreportlite]
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite2    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ee_payrollreportlite2]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ee_payrollreportlite2]
GO

/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eeapprovalspending]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eeapprovalspending]
GO

/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eedelegateformanagers]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eedelegateformanagers]
GO

/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eedetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eedetails]
GO

/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eedirectreports]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eedirectreports]
GO

/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eeelp]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eeelp]
GO

/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eeleaveperiods_by_type_for_year]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eeleaveperiods_by_type_for_year]
GO

/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eeleaverequests]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eeleaverequests]
GO

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eeleaverequests_adminview]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eeleaverequests_adminview]
GO

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eeleaverequests_overlapping]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eeleaverequests_overlapping]
GO

/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_eesearchbyname]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_eesearchbyname]
GO

/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_elpinstance]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_elpinstance]
GO

/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_elpinstance_elpreliefs]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_elpinstance_elpreliefs]
GO

/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_elprelief]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_elprelief]
GO

/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_leaveperiod]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_leaveperiod]
GO

/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_leavetype_by_id]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_leavetype_by_id]
GO

/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_leavetype_by_name]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_leavetype_by_name]
GO

/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_leavetypes_admin]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_leavetypes_admin]
GO

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_leavetypes_ee_other_leave]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_leavetypes_ee_other_leave]
GO

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_leavetypes_ee_requests]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_leavetypes_ee_requests]
GO

/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_publicholidays]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_publicholidays]
GO

/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_reject_cancel_request]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_reject_cancel_request]
GO

/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_reject_leave_request]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_reject_leave_request]
GO

/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_save_carryover]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_save_carryover]
GO

/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_save_elp_activation]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_save_elp_activation]
GO

/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_save_elprelief]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_save_elprelief]
GO

/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_save_leave_request_admin]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_save_leave_request_admin]
GO

/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_save_new_leave_request_standard]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_save_new_leave_request_standard]
GO

/****** Object:  Stored Procedure dbo.usp_save_user    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_save_user]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_save_user]
GO

/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_userdetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_userdetails]
GO

/****** Object:  View dbo.qryEEDetails    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryEEDetails]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryEEDetails]
GO

/****** Object:  View dbo.qryELPDetails    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryELPDetails]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryELPDetails]
GO

/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryLeavePeriodDetails]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryLeavePeriodDetails]
GO

/****** Object:  View dbo.qryPayrollReport    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryPayrollReport]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryPayrollReport]
GO

/****** Object:  View dbo.qryPayrollReport2    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryPayrollReport2]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryPayrollReport2]
GO

/****** Object:  View dbo.qryUserApprovalsAll    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryUserApprovalsAll]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryUserApprovalsAll]
GO

/****** Object:  View dbo.qryWorkerPrivateAll    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryWorkerPrivateAll]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryWorkerPrivateAll]
GO

/****** Object:  Table [dbo].[fkeys]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[fkeys]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[fkeys]
GO

/****** Object:  Table [dbo].[tblCarryOver]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblCarryOver]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblCarryOver]
GO

/****** Object:  Table [dbo].[tblELPInstance]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblELPInstance]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblELPInstance]
GO

/****** Object:  Table [dbo].[tblELPRelief]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblELPRelief]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblELPRelief]
GO

/****** Object:  Table [dbo].[tblEeType]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblEeType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblEeType]
GO

/****** Object:  Table [dbo].[tblLeavePeriod]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblLeavePeriod]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblLeavePeriod]
GO

/****** Object:  Table [dbo].[tblLeaveType]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblLeaveType]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblLeaveType]
GO

/****** Object:  Table [dbo].[tblPublicHoliday]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPublicHoliday]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPublicHoliday]
GO

/****** Object:  Table [dbo].[tblUser]    Script Date: 02/05/2002 14:23:41 ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblUser]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblUser]
GO

/****** Table WorkerPrivate. Added by [MFILLAST 08-2006] ******/
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WorkerPrivate]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[WorkerPrivate]
GO



/****** Object:  Table [dbo].[fkeys]    Script Date: 02/05/2002 14:23:50 ******/
CREATE TABLE [dbo].[fkeys] (
	[PKTABLE_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKTABLE_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN1_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN2_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN3_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN4_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN5_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN6_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN7_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN8_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN9_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN10_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN11_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN12_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN13_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN14_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN15_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FKCOLUMN16_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[FK_NAME] [varchar] (32) COLLATE SQL_Latin1_General_CP850_CI_AI NULL 
) ON [PRIMARY]
GO


/* table created by [MFILLAST 08-2006] : store a copy of cdis */
CREATE TABLE [dbo].[WorkerPrivate] (
[WWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NOT NULL  ,
[FirstName] [varchar] (30)  ,
[LastName] [varchar] (30)   ,
[ShortID] [char] (8)  ,
[StatCode] [char] (1)  ,
[DomainAddress] [varchar] (80),  
[EntityCode] [char] (3)  ,
[MailStop] [varchar] (50) , 
[MgrName] [varchar] (50)  ,
[MgrWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI,  
[RegionCode] [varchar] (4) ,
[SiteCode] [char] (2) ,
[TermDate] [smalldatetime] ,
[LastHireDate] [smalldatetime] ,
[EmpTypeCode] [char] (3) COLLATE SQL_Latin1_General_CP850_CI_AI NOT NULL ,
[OriginalHireDate] [smalldatetime]  ,
[SchedTypeCode] [varchar] (2) ,
[OrgUnitDescr] [varchar] (30) ,
[JobType] [varchar] (1),
[CdisCopyUpdateDate] [smalldatetime] 
)ON [PRIMARY]
GO
ALTER TABLE [dbo].[WorkerPrivate] WITH NOCHECK ADD 
	CONSTRAINT [PK_WorkerPrivate] PRIMARY KEY  NONCLUSTERED 
	(
		[WWID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO



/****** Object:  Table [dbo].[tblCarryOver]    Script Date: 02/05/2002 14:23:51 ******/
CREATE TABLE [dbo].[tblCarryOver] (
	[lngID] [int] IDENTITY (1, 1) NOT NULL ,
	[strEEWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NOT NULL ,
	[lngYear] [smallint] NOT NULL ,
	[lngDays] [real] NULL ,
	[strEnteredByWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[datEntered] [datetime] NULL ,
	[strComments] [varchar] (100) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[blnPreArranged] [bit] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblELPInstance]    Script Date: 02/05/2002 14:23:52 ******/
CREATE TABLE [dbo].[tblELPInstance] (
	[lngID] [int] IDENTITY (1, 1) NOT NULL ,
	[strEEWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[datActivated] [datetime] NULL ,
	[strActivatedByWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[dblInitialDaysBanked] [real] NULL ,
	[dblTargetDays] [real] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblELPRelief]    Script Date: 02/05/2002 14:23:53 ******/
CREATE TABLE [dbo].[tblELPRelief] (
	[lngID] [int] IDENTITY (1, 1) NOT NULL ,
	[lngELPID] [int] NULL ,
	[lngYear] [smallint] NULL ,
	[datEntered] [datetime] NULL ,
	[strEnteredByWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[strComments] [varchar] (50) COLLATE SQL_Latin1_General_CP850_CI_AI NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblEeType]    Script Date: 02/05/2002 14:23:54 ******/
CREATE TABLE [dbo].[tblEeType] (
	[lngID] [int] IDENTITY (1, 1) NOT NULL ,
	[strEmpTypeCode] [char] (3) COLLATE SQL_Latin1_General_CP850_CI_AI NOT NULL ,
	[strDescription] [varchar] (30) COLLATE SQL_Latin1_General_CP850_CI_AI NOT NULL ,
	[blnBlueBadge] [bit] NOT NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblLeavePeriod]    Script Date: 02/05/2002 14:23:54 ******/
CREATE TABLE [dbo].[tblLeavePeriod] (
	[lngID] [int] IDENTITY (1, 1) NOT NULL ,
	[strEEWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[strApproverWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[lngLeaveTypeID] [int] NULL ,
	[datStartDate] [smalldatetime] NULL ,
	[strStartTime] [char] (2) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[datEndDate] [smalldatetime] NULL ,
	[strEndTime] [char] (2) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[datRaised] [datetime] NOT NULL ,
	[datApproved] [datetime] NULL ,
	[datRejected] [datetime] NULL ,
	[datCancelRequested] [datetime] NULL ,
	[datCancelApproved] [datetime] NULL ,
	[datCancelRejected] [datetime] NULL ,
	[strRequestComments] [varchar] (100) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[strResponseComments] [varchar] (100) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[lngELPID] [int] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblLeaveType]    Script Date: 02/05/2002 14:23:55 ******/
CREATE TABLE [dbo].[tblLeaveType] (
	[lngID] [int] IDENTITY (1, 1) NOT NULL ,
	[strLeaveTypeName] [varchar] (30) COLLATE SQL_Latin1_General_CP850_CI_AI NOT NULL ,
	[blnEERequests] [bit] NULL ,
	[blnAdminRequests] [bit] NULL ,
	[blnRequestBeforeAccrued] [bit] NULL ,
	[dblMinimumDays] [real] NULL ,
	[dblEntitlement] [real] NULL ,
	[dblDaysBeforeStopsLegalAdjAccrual] [real] NULL ,
	[blnDaysBeforeStopsIsConsecutive] [bit] NULL ,
	[blnIsOtherLeave] [bit] NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblPublicHoliday]    Script Date: 02/05/2002 14:23:56 ******/
CREATE TABLE [dbo].[tblPublicHoliday] (
	[lngID] [int] IDENTITY (1, 1) NOT NULL ,
	[datDate] [datetime] NULL ,
	[strDescription] [varchar] (50) COLLATE SQL_Latin1_General_CP850_CI_AI NULL 
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[tblUser]    Script Date: 02/05/2002 14:23:56 ******/
CREATE TABLE [dbo].[tblUser] (
	[strWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NOT NULL ,
	[datDOB] [datetime] NULL ,
	[endDate] [datetime] NULL,
	[blnIsAdmin] [bit] NOT NULL ,
	[strDelegateWWID] [char] (8) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[blnIsException] [bit] NOT NULL ,
	[strExceptionComments] [varchar] (50) COLLATE SQL_Latin1_General_CP850_CI_AI NULL ,
	[datAquisitionDate] [smalldatetime] NULL ,
	[blnIsExemptStatusChanged] [bit] NULL ,
	[blnIsEELeaveTracked] [bit] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblCarryOver] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblCarryOver] PRIMARY KEY  CLUSTERED 
	(
		[lngID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE  CLUSTERED  INDEX [IX_tblPublicHoliday_datDate] ON [dbo].[tblPublicHoliday]([datDate]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblCarryOver] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblCarryOver_lngDays] DEFAULT (0) FOR [lngDays],
	CONSTRAINT [DF_tblCarryOverEOY_strEnteredByWWID] DEFAULT ('System') FOR [strEnteredByWWID],
	CONSTRAINT [DF_tblCarryOverEOY_datEntered] DEFAULT (getdate()) FOR [datEntered],
	CONSTRAINT [DF_tblCarryOver_blnPreArranged] DEFAULT (0) FOR [blnPreArranged]
GO

ALTER TABLE [dbo].[tblELPInstance] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblELPInstance] PRIMARY KEY  NONCLUSTERED 
	(
		[lngID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblELPRelief] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblELPRelief] PRIMARY KEY  NONCLUSTERED 
	(
		[lngID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_tblELPRelief_ELPID_Year] UNIQUE  NONCLUSTERED 
	(
		[lngELPID],
		[lngYear]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblEeType] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblEeType] PRIMARY KEY  NONCLUSTERED 
	(
		[lngID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLeavePeriod] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblLeavePeriod_datRaised] DEFAULT (getdate()) FOR [datRaised],
	CONSTRAINT [PK_tblLeavePeriod] PRIMARY KEY  NONCLUSTERED 
	(
		[lngID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblLeaveType] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblLeaveType_blnEERequests] DEFAULT (0) FOR [blnEERequests],
	CONSTRAINT [DF_tblLeaveType_blnAdminRequests] DEFAULT (0) FOR [blnAdminRequests],
	CONSTRAINT [DF_tblLeaveType_blnRequestBeforeAccrued] DEFAULT (0) FOR [blnRequestBeforeAccrued],
	CONSTRAINT [DF_tblLeaveType_lngMinimumDays] DEFAULT (0) FOR [dblMinimumDays],
	CONSTRAINT [DF_tblLeaveType_lngEntitlement] DEFAULT (0) FOR [dblEntitlement],
	CONSTRAINT [DF_tblLeaveType_lngDaysBeforeStopsLegalAdjAccrual] DEFAULT ((-1)) FOR [dblDaysBeforeStopsLegalAdjAccrual],
	CONSTRAINT [DF_tblLeaveType_blnAllDaysIncludedBeforeStopsLegalAdjAccrual] DEFAULT (0) FOR [blnDaysBeforeStopsIsConsecutive],
	CONSTRAINT [DF_tblLeaveType_blnIsOtherLeave] DEFAULT (1) FOR [blnIsOtherLeave],
	CONSTRAINT [PK_tblLeaveType] PRIMARY KEY  NONCLUSTERED 
	(
		[lngID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_tblLeaveType_strLeaveTypeName] UNIQUE  NONCLUSTERED 
	(
		[strLeaveTypeName]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPublicHoliday] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblPublicHoliday] PRIMARY KEY  NONCLUSTERED 
	(
		[lngID]
	)  ON [PRIMARY] 

GO

ALTER TABLE [dbo].[tblUser] WITH NOCHECK ADD 
	CONSTRAINT [DF_tblUser_blnIsAdmin] DEFAULT (0) FOR [blnIsAdmin],
	CONSTRAINT [DF_tblUser_blnIsException] DEFAULT (0) FOR [blnIsException],
	CONSTRAINT [DF_tblUser_blnExemptStatusChanged] DEFAULT (0) FOR [blnIsExemptStatusChanged],
	CONSTRAINT [PK_tblUser] PRIMARY KEY  NONCLUSTERED 
	(
		[strWWID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_tblCarryOver_datEntered] ON [dbo].[tblCarryOver]([datEntered]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblCarryOver_strEEWWID] ON [dbo].[tblCarryOver]([strEEWWID]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblCarryOver_lngYear] ON [dbo].[tblCarryOver]([lngYear]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblELPInstance_strEEWWID] ON [dbo].[tblELPInstance]([strEEWWID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblELPInstance_datActivated] ON [dbo].[tblELPInstance]([datActivated]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblELPRelief_lngELPID] ON [dbo].[tblELPRelief]([lngELPID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblELPRelief_lngYear] ON [dbo].[tblELPRelief]([lngYear]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblEeType_strEmpTypeCode] ON [dbo].[tblEeType]([strEmpTypeCode]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblLeavePeriod_datStartDate] ON [dbo].[tblLeavePeriod]([datStartDate]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblLeavePeriod_datEndDate] ON [dbo].[tblLeavePeriod]([datEndDate]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblLeavePeriod_strEEWWID] ON [dbo].[tblLeavePeriod]([strEEWWID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblLeavePeriod_strApproverWWID] ON [dbo].[tblLeavePeriod]([strApproverWWID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblLeavePeriod_lngELPID] ON [dbo].[tblLeavePeriod]([lngELPID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblLeavePeriod_StartTime] ON [dbo].[tblLeavePeriod]([strStartTime]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblLeavePeriod_EndTime] ON [dbo].[tblLeavePeriod]([strEndTime]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [xlplt1] ON [dbo].[tblLeavePeriod]([lngLeaveTypeID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

 CREATE  INDEX [IX_tblUser_DelegateWWID] ON [dbo].[tblUser]([strDelegateWWID]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblELPRelief] ADD 
	CONSTRAINT [FK_tblELPRelief_tblELPInstance] FOREIGN KEY 
	(
		[lngELPID]
	) REFERENCES [dbo].[tblELPInstance] (
		[lngID]
	)
GO

ALTER TABLE [dbo].[tblLeavePeriod] ADD 
	CONSTRAINT [FK_tblLeavePeriod_tblLeaveType] FOREIGN KEY 
	(
		[lngLeaveTypeID]
	) REFERENCES [dbo].[tblLeaveType] (
		[lngID]
	)
GO








SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.qryWorkerPrivateAll    Script Date: 02/05/2002 14:23:57 ******/

/****** Object:  View dbo.qryWorkerPrivateAll    Script Date: 18/04/2002 14:09:01 ******/

/****** Object:  View dbo.qryWorkerPrivateAll    Script Date: 25/02/2002 16:35:38 ******/

/****** Object:  View dbo.qryWorkerPrivateAll    Script Date: 01/06/2001 19:10:38 ******/
/*MODIF [MFILLAST 08-2006]
CREATE VIEW dbo.qryWorkerPrivateAll
AS
SELECT *
FROM dbn_cdisprivate01.dbo.WorkerPrivate
*/
GO
CREATE VIEW dbo.qryWorkerPrivateAll
AS
SELECT *
FROM dbo.WorkerPrivate
/* end of MODIF [MFILLAST 08-2006]*/

GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO







SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.qryEEDetails    Script Date: 02/05/2002 14:23:57 ******/




/****** Object:  View dbo.qryEEDetails    Script Date: 18/04/2002 14:09:01 ******/

CREATE VIEW dbo.qryEEDetails
AS
SELECT dbo.qryWorkerPrivateAll.*, 
    EvacUser.strWWID AS LocalWWID, EvacUser.*, 
    EEType.*
FROM dbo.qryWorkerPrivateAll LEFT OUTER JOIN
    dbo.tblUser EvacUser ON 
    dbo.qryWorkerPrivateAll.WWID = EvacUser.strWWID LEFT OUTER
     JOIN
    dbo.tblEeType EEType ON 
    dbo.qryWorkerPrivateAll.EmpTypeCode = EEType.strEmpTypeCode



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.qryELPDetails    Script Date: 02/05/2002 14:23:57 ******/

/****** Object:  View dbo.qryELPDetails    Script Date: 18/04/2002 14:09:01 ******/

/****** Object:  View dbo.qryELPDetails    Script Date: 25/02/2002 16:35:38 ******/

/****** Object:  View dbo.qryELPDetails    Script Date: 01/06/2001 19:10:38 ******/
CREATE VIEW dbo.qryELPDetails
AS
SELECT dbo.tblELPInstance.lngID AS ELPID, 
    dbo.tblLeavePeriod.lngID AS LeavePeriodID, 
    dbo.tblELPInstance.strEEWWID, 
    dbo.tblELPInstance.datActivated, 
    dbo.tblELPInstance.strActivatedByWWID, 
    dbo.tblELPInstance.dblInitialDaysBanked, 
    dbo.tblELPInstance.dblTargetDays, 
    dbo.tblLeavePeriod.strApproverWWID, 
    dbo.tblLeavePeriod.lngLeaveTypeID, 
    dbo.tblLeavePeriod.strStartTime, 
    dbo.tblLeavePeriod.datStartDate, 
    dbo.tblLeavePeriod.datEndDate, 
    dbo.tblLeavePeriod.strEndTime, 
    dbo.tblLeavePeriod.datRaised, 
    dbo.tblLeavePeriod.datApproved, 
    dbo.tblLeavePeriod.datRejected, 
    dbo.tblLeavePeriod.datCancelRequested, 
    dbo.tblLeavePeriod.datCancelApproved, 
    dbo.tblLeavePeriod.strRequestComments, 
    dbo.tblLeavePeriod.strResponseComments, 
    dbo.tblLeavePeriod.datCancelRejected
FROM dbo.tblELPInstance LEFT OUTER JOIN
    dbo.tblLeavePeriod ON 
    dbo.tblELPInstance.lngID = dbo.tblLeavePeriod.lngELPID
WHERE (dbo.tblLeavePeriod.datCancelApproved IS NULL) AND 
    (dbo.tblLeavePeriod.datRejected IS NULL)




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 02/05/2002 14:23:57 ******/

/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 18/04/2002 14:09:01 ******/

/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 25/02/2002 16:35:38 ******/

/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 01/06/2001 19:10:38 ******/
CREATE VIEW dbo.qryLeavePeriodDetails
AS
SELECT dbo.tblLeavePeriod.lngID, dbo.tblLeavePeriod.strEEWWID, 
    dbo.tblLeavePeriod.strApproverWWID, 
    dbo.tblLeavePeriod.lngLeaveTypeID, 
    dbo.tblLeavePeriod.datStartDate, 
    dbo.tblLeavePeriod.strStartTime, 
    dbo.tblLeavePeriod.datEndDate, 
    dbo.tblLeavePeriod.strEndTime, 
    dbo.tblLeavePeriod.datRaised, 
    dbo.tblLeavePeriod.datRejected, 
    dbo.tblLeavePeriod.datApproved, 
    dbo.tblLeavePeriod.datCancelRequested, 
    dbo.tblLeavePeriod.datCancelApproved,
    dbo.tblLeavePeriod.datCancelRejected,
    dbo.tblLeavePeriod.strRequestComments, 
    dbo.tblLeavePeriod.strResponseComments, 
    dbo.tblLeavePeriod.lngELPID, 
    dbo.tblLeaveType.strLeaveTypeName, 
    dbo.tblLeaveType.blnEERequests, 
    dbo.tblLeaveType.blnAdminRequests, 
    dbo.tblLeaveType.blnRequestBeforeAccrued, 
    dbo.tblLeaveType.dblMinimumDays, 
    dbo.tblLeaveType.dblEntitlement, 
    dbo.tblLeaveType.dblDaysBeforeStopsLegalAdjAccrual, 
    dbo.tblLeaveType.blnDaysBeforeStopsIsConsecutive
FROM dbo.tblLeavePeriod INNER JOIN
    dbo.tblLeaveType ON 
    dbo.tblLeaveType.lngID = dbo.tblLeavePeriod.lngLeaveTypeID





GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.qryPayrollReport    Script Date: 02/05/2002 14:23:57 ******/


/****** Object:  View dbo.qryPayrollReport    Script Date: 18/04/2002 14:09:01 ******/

CREATE VIEW dbo.qryPayrollReport AS

SELECT 	u.WWID, 
	u.LastName + ', ' + u.FirstName   AS FullName, 
	lv.strLeaveTypeName as LeaveType, 
	l.datStartDate as StartDate, 
	l.strStartTime as StartTime, 
	l.datEndDate as EndDate, 
	l.strEndTime as EndTime,
	l.datCancelApproved as CancelApproved,
	u.StatCode as Status,
	u.JobType as ExemptStatus,
	u.LastHireDate as ServiceDate,
	et.blnBlueBadge as IsBlueBadge,
	u.OriginalHireDate as ODOH,
	ee.datDOB as DOB

FROM qryEEDetails u LEFT OUTER JOIN tblLeavePeriod l 
			ON u.WWID = l.strEEWWID 
			LEFT OUTER JOIN tblLeaveType lv 
				ON l.lngLeaveTypeId = lv.lngID
				LEFT OUTER JOIN tblEeType et
					ON u.EmpTypeCode = et.strEmpTypeCode
					LEFT OUTER JOIN tblUser ee
					ON u.WWID = ee.strWWID
				
WHERE u.LocalWWID <> '' 



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.qryPayrollReport2    Script Date: 02/05/2002 14:23:57 ******/



/****** Object:  View dbo.qryPayrollReport2    Script Date: 18/04/2002 14:09:01 ******/


CREATE VIEW dbo.qryPayrollReport2 AS

SELECT 	u.WWID, /* Payroll Rpt , HR Rpt, Admin Rpt */
	u.LastName + ', ' + u.FirstName   AS FullName,  /* Payroll Rpt, HR Rpt */
	lv.strLeaveTypeName as LeaveType, /* Payroll Rpt */
	l.datStartDate as StartDate, /* Payroll Rpt */
	l.strStartTime as StartTime, /* Payroll Rpt */
	l.datEndDate as EndDate, /* Payroll Rpt */
	l.strEndTime as EndTime, /* Payroll Rpt */
	l.datRaised as LeaveRequestRaised, /* Payroll Rpt */
	l.datApproved as Approved, /* Payroll Rpt */
	l.datRejected as Rejected, /* Payroll Rpt */
	l.datCancelRequested as CancelRequested, /* Payroll Rpt */
	l.datCancelApproved as CancelApproved, /* Payroll Rpt */
	l.datCancelRejected as CancelRejected, /* Payroll Rpt */

	u.StatCode as Status, /* Payroll Rpt, HR Rpt */
	u.TermDate as TeminationDate,
	u.JobType as ExemptStatus, /* HR Rpt, Admin Rpt */
	u.LastHireDate as LDOH, /* Admin Rpt */
	u.OriginalHireDate as ODOH, /* Admin Rpt */
	et.blnBlueBadge as IsBlueBadge,
	u.blnIsAdmin as IsAdmin, /* Admin Rpt */
	u.blnIsEELeaveTracked as IsEELeaveTracked,  /* Admin Rpt */
	
	ee.datDOB as DOB, /* Admin Rpt */
	u.MgrName as ManagerName, /* HR Rpt */
	u.MgrWWID as ManagerWWID, /* HR Rpt */
	u.RegionCode as RegionCode, /* Payroll Rpt */
	u.SiteCode as SiteCode, /* Payroll Rpt */
	u.strDelegateWWID as DelegateWWID


FROM qryEEDetails u LEFT OUTER JOIN tblLeavePeriod l 
			ON u.WWID = l.strEEWWID 
			LEFT OUTER JOIN tblLeaveType lv 
				ON l.lngLeaveTypeId = lv.lngID
				LEFT OUTER JOIN tblEeType et
					ON u.EmpTypeCode = et.strEmpTypeCode
					LEFT OUTER JOIN tblUser ee
					ON u.WWID = ee.strWWID
				
WHERE u.LocalWWID <> '' 




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  View dbo.qryUserApprovalsAll    Script Date: 02/05/2002 14:23:57 ******/


/****** Object:  View dbo.qryUserApprovalsAll    Script Date: 18/04/2002 14:09:01 ******/
CREATE VIEW dbo.qryUserApprovalsAll
AS
SELECT
	EE.MgrWWID AS strEEManagerWWID,
	Manager.strDelegateWWID AS strManagerDelegateWWID,
	tblLeavePeriod.*
FROM (tblLeavePeriod LEFT JOIN
    qryWorkerPrivateAll AS EE ON 
    tblLeavePeriod.strEEWWID = EE.WWID) LEFT JOIN
    tblUser AS Manager ON EE.MgrWWID = Manager.strWWID






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_elp_data_for_employee    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_elp_data_for_employee    Script Date: 18/04/2002 14:09:02 ******/

/****** Object:  Stored Procedure dbo.pr_evc_elp_data_for_employee    Script Date: 25/02/2002 16:35:38 ******/

CREATE PROCEDURE dbo.pr_evc_elp_data_for_employee(@vWWID	AS CHAR(8))
AS

	SET NOCOUNT ON

	SELECT datActivated, dblInitialDaysBanked, dblTargetDays,lngID
	  FROM dbo.tblELPInstance
	 WHERE strEEWWID = @vWWID

	RETURN




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_employee_carry_over_for_year    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_employee_carry_over_for_year    Script Date: 18/04/2002 14:09:02 ******/

/****** Object:  Stored Procedure dbo.pr_evc_employee_carry_over_for_year    Script Date: 25/02/2002 16:35:39 ******/

CREATE PROCEDURE dbo.pr_evc_employee_carry_over_for_year(@vWWID	AS CHAR(8),
							@vYear	AS SMALLINT)
AS

	SET NOCOUNT ON

	SELECT SUM(lngDays) AS NumCarried
 	  FROM tblCarryOver
	 WHERE strEEWWID = @vWWID
	   AND lngYear = @vYear

	RETURN




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_employee_leave_for_year    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_employee_leave_for_year    Script Date: 18/04/2002 14:09:02 ******/

/****** Object:  Stored Procedure dbo.pr_evc_employee_leave_for_year    Script Date: 25/02/2002 16:35:39 ******/

CREATE PROCEDURE dbo.pr_evc_employee_leave_for_year(@vWWID 	AS CHAR(8),
							@vYear	AS SMALLINT)
AS

	SET NOCOUNT ON

	DECLARE @vStartDate	AS DATETIME,
		@vEndDate	AS DATETIME

	SELECT @vStartDate = CONVERT(DATETIME,'1/1/' + CAST(@vYear AS VARCHAR), 103)

	SELECT @vEndDate = CONVERT(DATETIME,'31/12/' + CAST(@vYear AS VARCHAR), 103)

	SELECT lp.datStartDate, lp.strStartTime, lp.datEndDate, lp.strEndTime, 
       	       lp.lngELPID, lt.dblDaysBeforeStopsLegalAdjAccrual, 
       	       lt.blnDaysBeforeStopsIsConsecutive, lt.blnIsOtherLeave
  	  FROM dbo.tblLeavePeriod lp INNER JOIN dbo.tblLeaveType lt 
				     ON lt.lngID = lp.lngLeaveTypeID
 	 WHERE lp.strEEWWID = @vWWID
           AND (lp.datEndDate >=  @vStartDate
           OR  lp.datStartdate <= @vEndDate)
           AND lp.datCancelApproved IS NULL
           AND lp.datRejected IS NULL

	RETURN




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/* add by [MFILLAST 08-2006] */
CREATE PROCEDURE dbo.pr_evc_shannon_sites
AS

	SET NOCOUNT ON

	SELECT DISTINCT wp.SiteCode 
	  FROM tblUser u INNER JOIN dbo.qryWorkerPrivateAll wp
	                         ON u.strWWID = wp.WWID
	 WHERE wp.EntityCode = 508
	ORDER BY wp.SiteCode ASC

	RETURN



GO

/* end [MFILLAST 08-2006] */

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_base_data    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_base_data    Script Date: 18/04/2002 14:09:02 ******/

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_base_data    Script Date: 25/02/2002 16:35:39 ******/


CREATE PROCEDURE dbo.pr_evc_hr_report_base_data
AS

	SET NOCOUNT ON

	SELECT wp.WWID, wp.StatCode, CASE WHEN wp.JobType = 'E' THEN 'E' ELSE 'N' END AS JobType, 
	       wp.LastName + ', ' + wp.FirstName AS FullName, wp.MgrName, wp.SiteCode, wp.OriginalHireDate,
  	       wp.LastHireDate, et.blnBlueBadge, u.datDOB,u.endDate,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						ON wp.WWID = u.strWWID 
					       INNER JOIN dbo.tblEeType et 
					        ON wp.EmpTypeCode = et.strEmpTypeCode
	 WHERE wp.EntityCode = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
        ORDER BY FullName ASC

	RETURN




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_employee_data    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_employee_data    Script Date: 18/04/2002 14:09:02 ******/

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_employee_data    Script Date: 25/02/2002 16:35:39 ******/

CREATE PROCEDURE dbo.pr_evc_hr_report_employee_data(@vWWID	AS CHAR(8))
AS

	SET NOCOUNT ON

	SELECT wp.WWID, wp.StatCode, CASE WHEN wp.JobType = 'E' THEN 'E' ELSE 'N' END AS JobType, 
	       wp.LastName + ', ' + wp.FirstName AS FullName, wp.MgrName, wp.SiteCode, wp.OriginalHireDate,
  	       wp.LastHireDate, et.blnBlueBadge, u.datDOB,u.endDate,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						ON wp.WWID = u.strWWID 
					       INNER JOIN dbo.tblEeType et 
					        ON wp.EmpTypeCode = et.strEmpTypeCode
	 WHERE wp.EntityCode = 508
	   AND wp.WWID = @vWWID
		
	RETURN




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_manager_reports    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_manager_reports    Script Date: 18/04/2002 14:09:02 ******/

/****** Object:  Stored Procedure dbo.pr_evc_hr_report_manager_reports    Script Date: 25/02/2002 16:35:39 ******/

CREATE PROCEDURE dbo.pr_evc_hr_report_manager_reports(@vWWID	AS CHAR(8))
AS

	SET NOCOUNT ON

	SELECT wp.WWID, wp.StatCode, CASE WHEN wp.JobType = 'E' THEN 'E' ELSE 'N' END AS JobType, 
	       wp.LastName + ', ' + wp.FirstName AS FullName, wp.MgrName, wp.SiteCode, wp.OriginalHireDate,
  	       wp.LastHireDate, et.blnBlueBadge, u.datDOB,u.endDate,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						ON wp.WWID = u.strWWID 
					       INNER JOIN dbo.tblEeType et 
					        ON wp.EmpTypeCode = et.strEmpTypeCode
	 WHERE wp.EntityCode = 508
	   AND wp.MgrWWID = @vWWID
		
	RETURN




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_site_only    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_site_only    Script Date: 18/04/2002 14:09:02 ******/

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_site_only    Script Date: 25/02/2002 16:35:39 ******/




/* return all requests with days in this period including rejected and cancelled requests*/
CREATE PROCEDURE dbo.pr_evc_payroll_report_filtered_site_only(@vStartDate 	DATETIME,
					                      @vEndDate		DATETIME,
					                      @vSite	        CHAR(2))
AS

	SET NOCOUNT ON

	SELECT wp.WWID, wp.LastName + ', ' + wp.FirstName AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.EntityCode = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
 /*[MFILLAST 08-2006] : we focus on the period, not on the request's state
            OR (l.datRaised BETWEEN @vStartDate AND @vEndDate)
            OR (l.datApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datRejected BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelRequested BETWEEN  @vStartDate AND @vEndDate)
	    OR (l.datCancelRejected BETWEEN @vStartDate AND @vEndDate))
   */        AND l.datStartDate <= @vEndDate
  	   AND wp.sitecode = @vSite
         ORDER BY wp.WWID ASC, Startdate ASC

	RETURN
  				



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_with_site    Script Date: 02/05/2002 14:23:58 ******/

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_with_site    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_with_site    Script Date: 25/02/2002 16:35:40 ******/

CREATE PROCEDURE dbo.pr_evc_payroll_report_filtered_with_site(@vStartDate 	DATETIME,
					                      @vEndDate		DATETIME,
					                      @vStatus	        VARCHAR(20),
					                      @vSite	        CHAR(2))
AS

	SET NOCOUNT ON

	IF @vStatus = 'active'
	BEGIN

	SELECT wp.WWID, wp.LastName + ', ' + wp.FirstName AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.EntityCode = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
 /*[MFILLAST 08-2006] : we focus on the period, not on the request's state
             OR (l.datRaised BETWEEN @vStartDate AND @vEndDate)
            OR (l.datApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datRejected BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelRequested BETWEEN  @vStartDate AND @vEndDate)
	    OR (l.datCancelRejected BETWEEN @vStartDate AND @vEndDate))
 */          AND l.datStartDate <= @vEndDate
	   AND wp.StatCode NOT IN ('R', 'T')
  	   AND wp.sitecode = @vSite
         ORDER BY wp.WWID ASC, Startdate ASC

	END
	ELSE
	BEGIN

	SELECT wp.WWID, wp.LastName + ', ' + wp.FirstName AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.EntityCode = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
 /*[MFILLAST 08-2006] : we focus on the period, not on the request's state            OR (l.datRaised BETWEEN @vStartDate AND @vEndDate)
            OR (l.datApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datRejected BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelRequested BETWEEN  @vStartDate AND @vEndDate)
	    OR (l.datCancelRejected BETWEEN @vStartDate AND @vEndDate))
 */          AND l.datStartDate <= @vEndDate
	   AND wp.StatCode IN ('R', 'T')
  	   AND wp.sitecode = @vSite
         ORDER BY wp.WWID ASC, Startdate ASC

        END

	RETURN
  				



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_without_site    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_without_site    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.pr_evc_payroll_report_filtered_without_site    Script Date: 25/02/2002 16:35:40 ******/

CREATE PROCEDURE dbo.pr_evc_payroll_report_filtered_without_site(@vStartDate 	DATETIME,
					                         @vEndDate	DATETIME,
					                         @vStatus	VARCHAR(20))
AS

	SET NOCOUNT ON

	IF @vStatus = 'active'
	BEGIN

	SELECT wp.WWID, wp.LastName + ', ' + wp.FirstName AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.EntityCode = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
  /*[MFILLAST 08-2006] : we focus on the period, not on the request's state
             OR (l.datRaised BETWEEN @vStartDate AND @vEndDate)
            OR (l.datApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datRejected BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelRequested BETWEEN  @vStartDate AND @vEndDate)
	    OR (l.datCancelRejected BETWEEN @vStartDate AND @vEndDate))
  */         AND l.datStartDate <= @vEndDate
	   AND wp.StatCode NOT IN ('R', 'T')
         ORDER BY wp.WWID ASC, Startdate ASC

	END
	ELSE
	BEGIN

	SELECT wp.WWID, wp.LastName + ', ' + wp.FirstName AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.OrgUnitDescr
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.EntityCode = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
      /*[MFILLAST 08-2006] : we focus on the period, not on the request's state
             OR (l.datRaised BETWEEN @vStartDate AND @vEndDate)
            OR (l.datApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datRejected BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelApproved BETWEEN @vStartDate AND @vEndDate)
            OR (l.datCancelRequested BETWEEN  @vStartDate AND @vEndDate)
	    OR (l.datCancelRejected BETWEEN @vStartDate AND @vEndDate))
     */      AND l.datStartDate <= @vEndDate
	   AND wp.StatCode IN ('R', 'T')
         ORDER BY wp.WWID ASC, Startdate ASC

        END

	RETURN
  				



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 25/02/2002 16:35:40 ******/

/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 01/06/2001 19:10:44 ******/
CREATE PROCEDURE dbo.usp_approve_cancel_request
@lngID integer,
@strResponseComments varchar(100) AS
UPDATE tblLeavePeriod
SET datCancelApproved = GetDate(),
strResponseComments = @strResponseComments
WHERE lngID = @lngID
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred approving the cancellation of the leave request record.'
	return(0)
end
else
begin
	   -- Return 1 to the calling program to indicate success.
	print 'The leave request has been cancel-approved.'
	return(1)
end




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 25/02/2002 16:35:40 ******/

/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 01/06/2001 19:10:39 ******/
CREATE PROCEDURE dbo.usp_approve_leave_request
@lngID integer,
@strResponseComments varchar(100) AS
UPDATE tblLeavePeriod
SET datApproved = GetDate(),
strResponseComments = @strResponseComments
WHERE lngID = @lngID
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred creating the new leave request (standard) record.'
	return(0)
end
else
begin
	   -- Return 0 to the calling program to indicate success.
	print 'The leave request (standard) has been updated.'
	return(1)
end






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 25/02/2002 16:35:40 ******/

/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 01/06/2001 19:10:44 ******/
CREATE PROCEDURE dbo.usp_cancel_leave_request
@lngID integer,
@strApproverWWID char(8),
@strRequestComments varchar(100),
@datCancelRaised smalldatetime,
@datCancelApproved smalldatetime,
@datCancelRejected smalldatetime
AS
UPDATE tblLeavePeriod
	SET strApproverWWID = @strApproverWWID,
		datCancelRequested = @datCancelRaised,
		datCancelApproved = @datCancelApproved,
		datCancelRejected = @datCancelRejected,
		strRequestComments = @strRequestComments,
		lngELPID = NULL
	WHERE lngID = @lngID

Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the 1 to the calling program to indicate failure.
	print 'An error occurred updating the leave request record.'
	return(1)
end
else
begin
	-- Return 0 to the calling program to indicate success.
	print 'The leave request record has been updated.'
	return(0)
end




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 25/02/2002 16:35:40 ******/

/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 01/06/2001 19:10:39 ******/
CREATE PROCEDURE dbo.usp_carryover_by_ID
@lngID integer
AS
SELECT * FROM tblcarryover WHERE lngID = @lngID






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 25/02/2002 16:35:40 ******/

/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 01/06/2001 19:10:39 ******/
CREATE PROCEDURE dbo.usp_carryover_by_WWID_Year
@strEEWWID char(8),
@lngYear integer
AS
SELECT * FROM tblcarryovereoy WHERE strEEWWID = @strEEWWID and lngYear = @lngYear



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 18/04/2002 14:09:03 ******/

/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 25/02/2002 16:35:41 ******/

/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE dbo.usp_carryovers_by_WWID
@strEEWWID char(8)
AS
SELECT * FROM tblcarryover WHERE strEEWWID = @strEEWWID
ORDER BY lngYear



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 25/02/2002 16:35:41 ******/

/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE dbo.usp_create_user
@strEEWWID char(8)
AS
insert into tblUser
	(strWWID
	)
Values (
	@strEEWWID
	)

Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred creating the new user record.'
	return(0)
end
else
begin
	   -- Return 0 to the calling program to indicate success.
	print 'The new user record has been created with default entries (except WWID which was set).'
	return(1)
end





GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 25/02/2002 16:35:41 ******/

/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE dbo.usp_delete_leave_request
@lngID integer
AS
DELETE FROM tblLeavePeriod
WHERE lngID = @lngID

Declare @lngError int
Select @lngError = @@ERROR
if @lngError <> 0
	--Return error code
	return @lngError
else
	--Return 0 to indicate success.
	return 0






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport    Script Date: 02/05/2002 14:23:59 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport    Script Date: 25/02/2002 16:35:41 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport    Script Date: 01/06/2001 19:10:48 ******/
CREATE PROCEDURE dbo.usp_ee_payrollreport
(
	@datRptStartDate smalldatetime,
	@datRptEndDate	smalldatetime
)
AS

SELECT WWID, FullName, LeaveType, StartDate, StartTime, EndDate, EndTime, Status, ExemptStatus, ServiceDate, IsBlueBadge, ODOH, DOB
FROM dbo.qryPayrollReport
WHERE ((StartDate IS NULL AND EndDate IS NULL) --EE with no leave 
OR ( StartDate>= @datRptStartDate AND EndDate <= @datRptEndDate))  --EE with leave between these dates

ORDER BY WWID asc, StartDate asc









GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport_old    Script Date: 02/05/2002 14:24:00 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport_old    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport_old    Script Date: 25/02/2002 16:35:41 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreport_old    Script Date: 01/06/2001 19:10:45 ******/
CREATE PROCEDURE dbo.usp_ee_payrollreport_old
AS
SELECT *
FROM qryEEDetails 
WHERE isnull(LocalWWID,'x') <> 'x'
ORDER BY strWWID







GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite    Script Date: 02/05/2002 14:24:00 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite    Script Date: 25/02/2002 16:35:41 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite    Script Date: 01/06/2001 19:10:48 ******/
CREATE PROCEDURE dbo.usp_ee_payrollreportlite
(
	@datRptStartDate smalldatetime,
	@datRptEndDate	smalldatetime
)
AS

SELECT WWID, FullName, LeaveType, StartDate, StartTime, EndDate, EndTime, CancelApproved
FROM dbo.qryPayrollReport
WHERE 
((StartDate IS NULL AND EndDate IS NULL) OR (StartDate>= @datRptStartDate)) 
ORDER BY WWID asc, StartDate asc




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite2    Script Date: 02/05/2002 14:24:00 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite2    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_ee_payrollreportlite2    Script Date: 25/02/2002 16:35:41 ******/

CREATE PROCEDURE dbo.usp_ee_payrollreportlite2
(
	@datRptStartDate smalldatetime,
	@datRptEndDate	smalldatetime
)
AS

	SELECT wp.WWID, wp.LastName + ', ' + wp.FirstName AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected
	  FROM dbn_cdisprivate01.dbo.WorkerPrivate wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.EntityCode = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @datRptStartDate AND @datRptEndDate)
	    OR (l.datEndDate BETWEEN @datRptStartDate AND @datRptEndDate)
            OR (l.datStartDate < @datRptStartDate AND l.datEndDate > @datRptEndDate)
            OR (l.datRaised BETWEEN @datRptStartDate AND @datRptEndDate)
            OR (l.datApproved BETWEEN @datRptStartDate AND @datRptEndDate)
            OR (l.datRejected BETWEEN @datRptStartDate AND @datRptEndDate)
            OR (l.datCancelApproved BETWEEN @datRptStartDate AND @datRptEndDate)
            OR (l.datCancelRequested BETWEEN  @datRptStartDate AND @datRptEndDate)
	    OR (l.datCancelRejected BETWEEN @datRptStartDate AND @datRptEndDate))
           AND l.datStartDate <= @datRptEndDate
         ORDER BY wp.WWID ASC, Startdate ASC


	RETURN




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 02/05/2002 14:24:00 ******/

/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE dbo.usp_eeapprovalspending
@strWWID char(8)
AS
SELECT
	*
FROM	qryUserApprovalsAll
WHERE
	strEEWWID <> @strWWID
	AND
	(
		strEEManagerWWID = @strWWID
		OR strApproverWWID = @strWWID
		OR strManagerDelegateWWID = @strWWID
	)





GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 02/05/2002 14:24:00 ******/

/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 18/04/2002 14:09:04 ******/

/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 01/06/2001 19:10:45 ******/
CREATE PROCEDURE dbo.usp_eedelegateformanagers
@strWWID char(8)
AS
SELECT *
FROM qryEEDetails
WHERE strDelegateWWID = @strWWID






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 02/05/2002 14:24:00 ******/

/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 01/06/2001 19:10:45 ******/
CREATE PROCEDURE dbo.usp_eedetails
@strWWID char(8)
AS
SELECT *
FROM qryEEDetails
WHERE WWID = @strWWID







GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 02/05/2002 14:24:00 ******/

/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 01/06/2001 19:10:45 ******/
CREATE PROCEDURE dbo.usp_eedirectreports
@strWWID char(8)
AS
SELECT *
FROM qryEEDetails 
WHERE MgrWWID = @strWWID











GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 02/05/2002 14:24:01 ******/

/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE dbo.usp_eeelp
@strWWID char(8)
AS
SELECT *
FROM qryELPDetails
WHERE strEEWWID = @strWWID






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 02/05/2002 14:24:01 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE dbo.usp_eeleaveperiods_by_type_for_year
@strWWID char(8),
@strLeaveTypeName char(30),
@strYear char(4)
AS
SELECT
	@strYear AS YearParam,
	*
FROM qryLeavePeriodDetails
WHERE strEEWWID = @strWWID
AND strLeaveTypeName = @strLeaveTypeName
AND datStartDate <= Convert(SmallDateTime,'31 December ' + @strYear)
AND datEndDate >= Convert(SmallDateTime, '1 January ' + @strYear)
ORDER BY datStartDate





GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 02/05/2002 14:24:01 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE dbo.usp_eeleaverequests
@strWWID char(8),
@datStartDate smalldatetime,
@datEndDate smalldatetime
AS
Select * From qryLeavePeriodDetails
WHERE strEEWWID = @strWWID
AND 
(
	(
		datStartDate <= @datEndDate
 		AND datEndDate >= @datStartDate
	)
	--OR --**Modified to not show pending approvals
	--(
		--datApproved Is Null
		--AND datRejected Is Null
		--AND datCancelRequested Is Null
		--AND datCancelApproved Is Null
		--AND datCancelRejected Is Null
	--)
)
ORDER BY datStartDate, strStartTime, datEndDate, strEndTime











GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 02/05/2002 14:24:01 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 25/02/2002 16:35:42 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE dbo.usp_eeleaverequests_adminview
@strWWID char(8)
AS
Select * From qryLeavePeriodDetails
WHERE strEEWWID = @strWWID
ORDER BY datStartDate, strStartTime, datEndDate, strEndTime




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 02/05/2002 14:24:01 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 25/02/2002 16:35:43 ******/

/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 01/06/2001 19:10:50 ******/
CREATE PROCEDURE dbo.usp_eeleaverequests_overlapping
@strWWID char(8),
@datStartDate smalldatetime,
@datEndDate smalldatetime
AS
Select * From qryLeavePeriodDetails
WHERE strEEWWID = @strWWID
AND datStartDate <= @datEndDate
AND datEndDate >= @datStartDate
ORDER BY datStartDate, strStartTime, datEndDate, strEndTime









GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 02/05/2002 14:24:01 ******/

/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 25/02/2002 16:35:43 ******/

/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE dbo.usp_eesearchbyname
@strName varchar(50),
@lngMaxResults integer
AS
Set @strName = RTrim(@strName)
Set @lngMaxResults = @lngMaxResults + 1
SET ROWCOUNT @lngMaxResults
SELECT WWID, LastName, FirstName, MgrName, MailStop FROM qryEEDetails
WHERE LastName+","+FirstName Like(@strName + '%')
ORDER BY LastName, FirstName
Return @@rowcount




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 02/05/2002 14:24:01 ******/

/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 18/04/2002 14:09:05 ******/

/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 25/02/2002 16:35:43 ******/

/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 01/06/2001 19:10:50 ******/
CREATE PROCEDURE dbo.usp_elpinstance
@lngID integer
AS
SELECT *
FROM qryELPDetails
WHERE ELPID = @lngID







GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 25/02/2002 16:35:43 ******/

/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE dbo.usp_elpinstance_elpreliefs
@lngELPID integer
AS
SELECT *
FROM dbo.tblELPRelief 
WHERE lngELPID = @lngELPID
ORDER BY lngYear







GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 25/02/2002 16:35:43 ******/

/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE dbo.usp_elprelief
@lngID integer
AS
	SELECT *
	FROM dbo.tblELPRelief 
	WHERE lngID = @lngID




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 25/02/2002 16:35:43 ******/

/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE dbo.usp_leaveperiod
@lngID integer
AS
SELECT *
FROM dbo.tblLeavePeriod WHERE lngID = @lngID







GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 25/02/2002 16:35:43 ******/

/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 01/06/2001 19:10:41 ******/
CREATE PROCEDURE dbo.usp_leavetype_by_id
@lngID integer
AS
	SELECT *
	FROM dbo.tblLeaveType
	WHERE lngID = @lngID








GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 01/06/2001 19:10:41 ******/
CREATE PROCEDURE dbo.usp_leavetype_by_name
@strName char(30)
AS
	SELECT *
	FROM dbo.tblLeaveType
	WHERE strLeaveTypeName = @strName








GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE dbo.usp_leavetypes_admin AS
SELECT dbo.tblLeaveType.*
FROM dbo.tblLeaveType
ORDER BY strLeaveTypeName





GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE dbo.usp_leavetypes_ee_other_leave AS
SELECT dbo.tblLeaveType.*
FROM dbo.tblLeaveType
WHERE blnIsOtherLeave = 1
ORDER BY strLeaveTypeName




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE dbo.usp_leavetypes_ee_requests AS
SELECT dbo.tblLeaveType.*
FROM dbo.tblLeaveType
WHERE blnEERequests = 1






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 18/04/2002 14:09:06 ******/

/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 01/06/2001 19:10:43 ******/
CREATE PROCEDURE dbo.usp_publicholidays
AS
Select * From tblPublicHoliday
ORDER BY datDate





GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 02/05/2002 14:24:02 ******/

/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE dbo.usp_reject_cancel_request
@lngID integer,
@strResponseComments varchar(100) AS
UPDATE tblLeavePeriod
SET datCancelRejected = GetDate(),
strResponseComments = @strResponseComments
WHERE lngID = @lngID
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return 0 to the calling program to indicate failure.
	print 'An error occurred cancelling the leave request record.'
	return(0)
end
else
begin
	   -- Return 1  to the calling program to indicate success.
	print 'The leave request has been updated.'
	return(1)
end




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 01/06/2001 19:10:41 ******/
CREATE PROCEDURE dbo.usp_reject_leave_request
@lngID integer,
@strResponseComments varchar(100) AS
UPDATE tblLeavePeriod
SET datRejected = GetDate(),
strResponseComments = @strResponseComments
WHERE lngID = @lngID
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred creating the new leave request (standard) record.'
	return(0)
end
else
begin
	   -- Return 0 to the calling program to indicate success.
	print 'The leave request (standard) has been updated.'
	return(1)
end






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 25/02/2002 16:35:44 ******/

/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 01/06/2001 19:10:43 ******/
CREATE PROCEDURE dbo.usp_save_carryover
@strEEWWID char(8),
@lngYear integer,
@lngDays real,
@strEnteredByWWID char(8),
@datEntered smalldatetime,
@strComments varchar(100),
@blnPreArranged bit
AS
SELECT * FROM tblCarryOver WHERE
	strEEWWID = @strEEWWID AND
	lngYear = @lngYear AND
	blnPreArranged = @blnPreArranged

If @@ROWCOUNT = 0
	begin
		INSERT INTO tblCarryOver
			(strEEWWID,
			lngYear,
			lngDays,
			strEnteredByWWID,
			datEntered,
			strComments,
			blnPreArranged)
		VALUES
			(@strEEWWID,
			@lngYear,
			@lngDays,
			@strEnteredByWWID,
			@datEntered,
			@strComments,
			@blnPreArranged)
	end
else
	begin
		UPDATE tblCarryOver
			Set lngDays = @lngDays,
			datEntered = @datEntered,
			strComments = @strComments
		WHERE
			strEEWWID = @strEEWWID AND
			lngYear = @lngYear AND
			blnPreArranged = @blnPreArranged
	end

Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
	begin
		-- Return the error code to the calling program to indicate failure.
		print 'An error occurred updating the user record.'
		return(0)
	end
else
	begin
		   -- Return 0 to the calling program to indicate success.
		print 'The new user record has been updated.'
		return(1)
	end




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 25/02/2002 16:35:45 ******/

/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 01/06/2001 19:10:43 ******/
CREATE PROCEDURE dbo.usp_save_elp_activation
@lngELPID integer,
@strEEWWID char(8),
@datActivated smalldatetime,
@strActivatedBy char(8),
@dblInitialDaysBanked real,
@dblTargetDays real
AS
if @lngELPID = 0
	begin
		insert into tblELPInstance
			(strEEWWID,
			datActivated,
			strActivatedByWWID,
			dblInitialDaysBanked,
			dblTargetDays
			)
		Values (
			@strEEWWID,
			@datActivated,
			@strActivatedBy,
			@dblInitialDaysBanked,
			@dblTargetDays
			)
	end
else
	begin
		update tblELPInstance
			Set strEEWWID = @strEEWWID,
			datActivated = @datActivated,
			strActivatedByWWID = @strActivatedBy,
			dblInitialDaysBanked = @dblInitialDaysBanked,
			dblTargetDays = @dblTargetDays
		Where lngID = @lngELPID
	end
	

Declare @lngError int, @lngIdentity int
Select @lngError = @@ERROR
Select @lngIdentity = @@IDENTITY

If @lngELPID = 0
	begin
		If @lngError <> 0
			begin
				-- Return the 0 to the calling program to indicate failure.
				print 'An error occurred creating the ELP Instance.'
				return(0)
			end
		else
			begin
				   -- Return the id to the calling program to indicate success.
				print 'The ELP Instance was created successfully.'
				return(@lngIdentity)
			end
	end
else
	begin
		if @lngError <> 0
				--Return 0 to indicate error.
			return 0
		else
				--Return 1 to indicate success.
			return 1
	end
				




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 25/02/2002 16:35:45 ******/

/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE dbo.usp_save_elprelief
@blnAdd bit,
@lngYear integer,
@lngELPID integer,
@datDateEntered smalldatetime,
@strEnteredByWWID char(8)
AS
If @blnAdd = 1
	begin
		insert into tblELPRelief
			(lngYear,
			lngELPID,
			datEntered,
			strEnteredByWWID
			)
		Values (
			@lngYear,
			@lngELPID,
			@datDateEntered,
			@strEnteredByWWID
			)
	end
else
	begin
		delete from tblELPRelief
			where lngYear = @lngYear and
				lngELPID = @lngELPID
	end
	
Declare @lngError int
Select @lngError = @@ERROR

If @lngError <> 0
	begin
		-- Return the 0 to the calling program to indicate failure.
		print 'An error occurred updating the elp relief (admin) record.'
		return(0)
	end
else
	begin
		   -- Return 1 to the calling program to indicate success.
		print 'The elp relief was updated successfully.'
		return(1)
	end

GO
/* return 1 if there is a relief for this elp and this year, 0 else*/
CREATE PROCEDURE dbo.get_elprelief_year
@lngYear smallint,
@lngELPID integer,
AS
SELECT COUNT(*) as isRelief
FROM dbo.tblELPRelief
WHERE @lngYear = lngYear
AND @lngELPID = lngELPID
return
GO
/* return the number of reliefs for this elp (used to calcul then number of years to run) */
CREATE PROCEDURE dbo.get_nb_elprelief
@lngELPID integer,
AS
SELECT COUNT(*) as nbRelief
FROM dbo.tblELPRelief
WHERE @lngELPID = lngELPID
return



GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 25/02/2002 16:35:45 ******/

/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE dbo.usp_save_leave_request_admin
@lngID integer,
@strEEWWID char(8),
@strApproverWWID char(8),
@lngLeaveTypeID integer,
@datStartDate smalldatetime,
@strStartTime char(2),
@datEndDate smalldatetime,
@strEndTime char(2),
@strRequestComments varchar(100),
@strResponseComments varchar(100),
@lngELPID integer
 AS

If @lngID = 0
	begin
		insert into tblLeavePeriod
			(strEEWWID,
			strApproverWWID,
			lngLeaveTypeID,
			datStartDate,
			strStartTime,
			datEndDate,
			strEndTime,
			strRequestComments,
			strResponseComments,
			datApproved,
			lngELPID
			)
		Values (
			@strEEWWID,
			@strApproverWWID,
			@lngLeaveTypeID,
			@datStartDate,
			@strStartTime,
			@datEndDate,
			@strEndTime,
			@strRequestComments,
			@strResponseComments,
			GetDate(),
			@lngELPID
			)
	end
else
	begin
		update tblLeavePeriod
			Set strEEWWID = @strEEWWID,
			strApproverWWID = @strApproverWWID,
			lngLeaveTypeID = @lngLeaveTypeID,
			datStartDate = @datStartDate,
			strStartTime = @strStartTime,
			datEndDate = @datEndDate,
			strEndTime = @strEndTime,
			strRequestComments = @strRequestComments,
			strResponseComments = @strResponseComments,
			datApproved = GetDate(),
			lngELPID = @lngELPID
		where lngID = @lngID
	end

Declare @lngError int, @lngIdentity int
Select @lngError = @@ERROR
Select @lngIdentity = @@IDENTITY

if @lngID = 0
	begin
		If @lngError <> 0
			begin
				-- Return the 0 to the calling program to indicate failure.
				print 'An error occurred updating the leave period (admin) record.'
				return(0)
			end
		else
			begin
				   -- Return the id to the calling program to indicate success.
				print 'The leave period (admin) was created or updated successfully.'
				return(@lngIdentity)
			end
	end
else
	begin
		if @lngError <> 0
				--Return 0 to indicate error.
			return 0
		else
				--Return 1 to indicate success.
			return 1
	end






GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 25/02/2002 16:35:45 ******/

/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE dbo.usp_save_new_leave_request_standard
@strEEWWID char(8),
@strApproverWWID char(8),
@lngLeaveTypeID integer,
@datStartDate smalldatetime,
@strStartTime char(2),
@datEndDate smalldatetime,
@strEndTime char(2),
@strRequestComments varchar(100),
@lngELPID integer
 AS
insert into tblLeavePeriod
	(strEEWWID,
	strApproverWWID,
	lngLeaveTypeID,
	datStartDate,
	strStartTime,
	datEndDate,
	strEndTime,
	strRequestComments,
	lngELPID
	)
Values (
	@strEEWWID,
	@strApproverWWID,
	@lngLeaveTypeID,
	@datStartDate,
	@strStartTime,
	@datEndDate,
	@strEndTime,
	@strRequestComments,
	@lngELPID
	)

Declare @lngError int, @lngIdentity int
Select @lngError = @@ERROR
Select @lngIdentity = @@IDENTITY
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred creating the new leave request (standard) record.'
	return(0)
end
else
begin
	   -- Return 0 to the calling program to indicate success.
	print 'The new leave request (standard) has been created.'
	return(@lngIdentity)
end



	




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_save_user    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_save_user    Script Date: 18/04/2002 14:09:07 ******/

CREATE PROCEDURE dbo.usp_save_user
@strEEWWID char(8),
@strDelegateWWID char(8),
@datDOB smalldatetime,
@endDate smalldatetime,
@blnIsEELeaveTracked bit,
@blnIsAdmin bit
AS

IF (@strDelegateWWID IS NOT NULL)
BEGIN
  IF (LEN(LTRIM(@strDelegateWWID)) = 0)
  BEGIN
    SELECT @strDelegateWWID = NULL
  END
END

UPDATE tblUser
SET strDelegateWWID = @strDelegateWWID,
	datDOB = @datDOB,
	endDate = @endDate,
	blnIsEELeaveTracked = @blnIsEELeaveTracked,
	blnIsAdmin = @blnIsAdmin
WHERE 
	strWWID = @strEEWWID

Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred updating the user record.'
	return(0)
end
else
begin
	   -- Return 0 to the calling program to indicate success.
	print 'The new user record has been updated.'
	return(1)
end




GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 02/05/2002 14:24:03 ******/

/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 18/04/2002 14:09:07 ******/

/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 25/02/2002 16:35:45 ******/

/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 01/06/2001 19:10:48 ******/
CREATE PROCEDURE dbo.usp_userdetails
@strShortID char(8)
AS
SELECT *
FROM qryEEDetails
WHERE ShortID = @strShortID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
/* added by [MFILLAST 08-2006] : create a new entry in the cdis copy table */

CREATE PROCEDURE dbo.usp_insert_cdis_copy
@WWID char(8)  ,
@FirstName varchar(30) ,
@LastName varchar(30) , 
@ShortID char(8) ,
@StatCode char(1) ,
@DomainAddress varchar(80) ,
@EntityCode char(3) ,
@MailStop varchar(50) ,
@MgrName varchar(50) ,
@MgrWWID char(8) ,
@RegionCode varchar(4) ,
@SiteCode char(2) ,
@TermDate smalldatetime ,
@LastHireDate smalldatetime ,
@EmpTypeCode char(3) ,
@OriginalHireDate smalldatetime  ,
@SchedTypeCode varchar(2) ,
@OrgUnitDescr varchar(30),
@JobType varchar(1) AS
SET NOCOUNT ON
INSERT INTO WorkerPrivate
( WWID ,FirstName ,LastName , ShortID ,StatCode ,DomainAddress ,EntityCode ,MailStop ,MgrName ,MgrWWID ,RegionCode ,SiteCode ,TermDate ,LastHireDate ,EmpTypeCode ,OriginalHireDate ,SchedTypeCode ,JobType,OrgUnitDescr,CdisCopyUpdateDate )
VALUES (@WWID ,@FirstName ,@LastName ,@ShortID ,@StatCode ,@DomainAddress ,@EntityCode ,@MailStop ,@MgrName ,@MgrWWID ,@RegionCode ,@SiteCode ,@TermDate ,@LastHireDate ,@EmpTypeCode ,@OriginalHireDate ,@SchedTypeCode ,@JobType,@OrgUnitDescr,GetDate() )
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred inserting cdis copy.'
	return(0)
end
else
begin
	   -- Return 1 to the calling program to indicate success.
	print 'inserted in local copy of cdis .'
	return(1)
end


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
/* added by [MFILLAST 08-2006] : update the cdis copy table */
CREATE PROCEDURE dbo.usp_update_cdis_copy
@WWID char(8)  ,
@FirstName varchar(30) ,
@LastName varchar(30) , 
@ShortID char(8) ,
@StatCode char(1) ,
@DomainAddress varchar(80) ,
@EntityCode char(3) ,
@MailStop varchar(50) ,
@MgrName varchar(50) ,
@MgrWWID char(8) ,
@RegionCode varchar(4) ,
@SiteCode char(2) ,
@TermDate smalldatetime ,
@LastHireDate smalldatetime ,
@EmpTypeCode char(3) ,
@OriginalHireDate smalldatetime  ,
@SchedTypeCode varchar(2) ,
@OrgUnitDescr varchar(30),
@JobType varchar(1) AS
SET NOCOUNT ON
UPDATE WorkerPrivate
SET 
ShortID = @ShortID  ,
FirstName = @FirstName  ,
LastName = @LastName  , 
StatCode = @StatCode  ,
DomainAddress = @DomainAddress  ,
EntityCode = @EntityCode  ,
MailStop = @MailStop  ,
MgrName = @MgrName  ,
MgrWWID = @MgrWWID  ,
RegionCode = @RegionCode  ,
SiteCode = @SiteCode  ,
TermDate = @TermDate  ,
LastHireDate = @LastHireDate  ,
EmpTypeCode = @EmpTypeCode  ,
OriginalHireDate = @OriginalHireDate   ,
SchedTypeCode = @SchedTypeCode  ,
JobType = @JobType  ,
OrgUnitDescr =@OrgUnitDescr,
CdisCopyUpdateDate = GetDate()
WHERE WWID = @WWID
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred updating cdis copy.'
	return(0)
end
else
begin
	   -- Return 1 to the calling program to indicate success.
	print 'Local copy of cdis updated.'
	return(1)
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
/* added by [MFILLAST 08-2006] : return the date of the last update of info stored in the copy of cdis(using user's wwid). */
CREATE PROCEDURE dbo.usp_getLastUpdateCdisCopy_from_wwid
@strWWID char(8)
AS
SET NOCOUNT ON
SELECT CdisCopyUpdateDate
FROM WorkerPrivate
WHERE WWID = @strWWID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
/* added by [MFILLAST 08-2006] : return the date of the last update of info stored in the copy of cdis(using user's shortid). */
CREATE PROCEDURE dbo.usp_getLastUpdateCdisCopy_from_shortid
@strShortID char(8)
AS
SET NOCOUNT ON
SELECT CdisCopyUpdateDate
FROM WorkerPrivate
WHERE ShortID = @strShortID







GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

