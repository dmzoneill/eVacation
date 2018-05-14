USE [wds_evacation_dev]
GO
/****** Object:  User [absence_booking]    Script Date: 4/19/2018 9:29:56 AM ******/
CREATE USER [absence_booking] WITHOUT LOGIN WITH DEFAULT_SCHEMA=[absence_booking]
GO
/****** Object:  User [evac_dev]    Script Date: 4/19/2018 9:29:56 AM ******/
CREATE USER [evac_dev] WITHOUT LOGIN WITH DEFAULT_SCHEMA=[evac_dev]
GO
/****** Object:  User [EVAC-ADMIN]    Script Date: 4/19/2018 9:29:56 AM ******/
CREATE USER [EVAC-ADMIN] WITHOUT LOGIN WITH DEFAULT_SCHEMA=[dbo]
GO
/****** Object:  User [lab_bfadmin]    Script Date: 4/19/2018 9:29:56 AM ******/
CREATE USER [lab_bfadmin] FOR LOGIN [GER\lab_bfadmin] WITH DEFAULT_SCHEMA=[dbo]
GO
ALTER ROLE [db_owner] ADD MEMBER [absence_booking]
GO
ALTER ROLE [db_accessadmin] ADD MEMBER [absence_booking]
GO
ALTER ROLE [db_securityadmin] ADD MEMBER [absence_booking]
GO
ALTER ROLE [db_ddladmin] ADD MEMBER [absence_booking]
GO
ALTER ROLE [db_backupoperator] ADD MEMBER [absence_booking]
GO
ALTER ROLE [db_datareader] ADD MEMBER [absence_booking]
GO
ALTER ROLE [db_datawriter] ADD MEMBER [absence_booking]
GO
ALTER ROLE [db_owner] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_accessadmin] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_securityadmin] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_ddladmin] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_backupoperator] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_datareader] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_datawriter] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_denydatareader] ADD MEMBER [lab_bfadmin]
GO
ALTER ROLE [db_denydatawriter] ADD MEMBER [lab_bfadmin]
GO
/****** Object:  Schema [absence_booking]    Script Date: 4/19/2018 9:29:56 AM ******/
CREATE SCHEMA [absence_booking]
GO
/****** Object:  Schema [evac_dev]    Script Date: 4/19/2018 9:29:56 AM ******/
CREATE SCHEMA [evac_dev]
GO
/****** Object:  Schema [lab_bfadmin]    Script Date: 4/19/2018 9:29:56 AM ******/
CREATE SCHEMA [lab_bfadmin]
GO
/****** Object:  StoredProcedure [dbo].[clean_up_db]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[clean_up_db] AS
DELETE FROM [dbo].[WorkerPrivate]
      WHERE [EmployeeStatusCd] = 'T'


DELETE FROM [dbo].[WorkerPrivate]
      WHERE [CorporateEmailTxt] is null


DELETE FROM [dbo].[WorkerPrivate]
      WHERE [EndDt] < SYSDATETIME()


DELETE FROM [dbo].[WorkerPrivate]
      WHERE [EmployeeStatusCd] is null


GO
/****** Object:  StoredProcedure [dbo].[get_elprelief_year]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[get_elprelief_year]
@lngYear smallint,
@lngELPID integer
AS
SELECT COUNT(*) AS isRelief
FROM dbo.tblELPRelief
WHERE @lngYear = lngYear
AND @lngELPID = lngELPID


GO
/****** Object:  StoredProcedure [dbo].[get_emps]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[get_emps](@vWWID AS CHAR(8))
AS

	SET NOCOUNT ON

Select *
From WorkerPrivate
Where NextLevelWWID = @vWWID


GO
/****** Object:  StoredProcedure [dbo].[get_nb_elprelief]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[get_nb_elprelief]
@lngELPID integer
AS
SELECT COUNT(*) as nbRelief
FROM dbo.tblELPRelief
WHERE @lngELPID = lngELPID
return


GO
/****** Object:  StoredProcedure [dbo].[getWWID]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[getWWID] AS
SELECT WWID
  FROM [wds_evacation_dev].[dbo].[WorkerPrivate] 
return


GO

/****** Object:  StoredProcedure [dbo].[hol_cal_display]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[hol_cal_display](@vWWID AS CHAR(8))
AS

	SET NOCOUNT ON

Select Distinct wp.FirstNm, wp.LastNm, lp.datStartDate, lp.datEndDate, wp.WWID, lp.datApproved, lp.datCancelApproved, lp.shareLeaveWithTeamCalendar
From tblLeavePeriod lp INNER JOIN WorkerPrivate wp ON lp.strEEWWID = wp.WWID
Where wp.NextLevelWWID = @vWWID


GO
/****** Object:  StoredProcedure [dbo].[leave_type_add]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[leave_type_add](@vName AS VARCHAR(50),@vEERequests AS INTEGER,@vAdminRequests AS INTEGER,@vRequestsBeforeOccrued AS INTEGER,@vMinDays AS VARCHAR,@vEntitlement AS VARCHAR,@vDaysBeforeStopsLegalAdjAccrual AS INTEGER,@vDaysBeforeStopsIsConsecutive AS INTEGER,@vIsOtherLeave AS INTEGER)
as

set nocount on

insert into tblLeaveType
Values(@vName,@vEERequests,@vAdminRequests,@vRequestsBeforeOccrued,@vMinDays,@vEntitlement,@vDaysBeforeStopsLegalAdjAccrual,@vDaysBeforeStopsIsConsecutive,@vIsOtherLeave) 


GO
/****** Object:  StoredProcedure [dbo].[leave_type_delete]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[leave_type_delete](@vName AS VARCHAR(50))
as

set nocount on

delete from tblLeaveType where strLeaveTypeName = @vName


GO
/****** Object:  StoredProcedure [dbo].[leave_type_display]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[leave_type_display]
as

set nocount on

select strLeaveTypeName from tblLeaveType


GO
/****** Object:  StoredProcedure [dbo].[mail_cal_display]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[mail_cal_display](@vWWID AS CHAR(8),@vMyStartDate DATETIME,@vMyEndDate DATETIME)
as

set nocount on

select wp.FirstNm, wp.LastNm, lp.datStartDate, lp.datEndDate, wp.WWID, lp.datApproved, lp.datCancelApproved from tblLeavePeriod lp INNER JOIN WorkerPrivate wp on lp.strEEWWID = wp.WWID where lp.datStartDate BETWEEN @vMyStartDate AND @vMyEndDate AND lp.datEndDate BETWEEN @vMyStartDate AND @vMyEndDate AND wp.NextLevelWWID = @vWWID


GO
/****** Object:  StoredProcedure [dbo].[mans_man]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[mans_man](@vWWID AS Char(8))
as

set nocount on

select NextLevelWWID from WorkerPrivate
Where WWID = @vWWID AND EmployeeStatusCd = 'A'


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_elp_data_for_employee]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[pr_evc_elp_data_for_employee](@vWWID	AS CHAR(8))
AS
	SET NOCOUNT ON
	SELECT datActivated, dblInitialDaysBanked, dblTargetDays,lngID
	  FROM dbo.tblELPInstance
	 WHERE strEEWWID = @vWWID
	RETURN


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_employee_carry_over_for_year]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.pr_evc_employee_carry_over_for_year    Script Date: 02/05/2002 14:23:58 ******/
/****** Object:  Stored Procedure dbo.pr_evc_employee_carry_over_for_year    Script Date: 18/04/2002 14:09:02 ******/
/****** Object:  Stored Procedure dbo.pr_evc_employee_carry_over_for_year    Script Date: 25/02/2002 16:35:39 ******/
CREATE PROCEDURE [dbo].[pr_evc_employee_carry_over_for_year](@vWWID	AS CHAR(8),
							@vYear	AS SMALLINT)
AS
	SET NOCOUNT ON
	SELECT SUM(lngDays) AS NumCarried
 	  FROM tblCarryOver
	 WHERE strEEWWID = @vWWID
	   AND lngYear = @vYear
	RETURN


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_employee_leave_for_year]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Stored Procedure dbo.pr_evc_employee_leave_for_year    Script Date: 02/05/2002 14:23:58 ******/
/****** Object:  Stored Procedure dbo.pr_evc_employee_leave_for_year    Script Date: 18/04/2002 14:09:02 ******/
/****** Object:  Stored Procedure dbo.pr_evc_employee_leave_for_year    Script Date: 25/02/2002 16:35:39 ******/
CREATE PROCEDURE [dbo].[pr_evc_employee_leave_for_year](@vWWID 	AS CHAR(8),
							@vYear	AS SMALLINT)
AS
	SET NOCOUNT ON
	DECLARE @vStartDate	AS DATETIME,
		@vEndDate	AS DATETIME
	SELECT @vStartDate = CONVERT(DATETIME,'1/1/' + CAST(@vYear AS VARCHAR), 103)
	SELECT @vEndDate = CONVERT(DATETIME,'31/12/' + CAST(@vYear AS VARCHAR), 103)
	SELECT lp.datStartDate, lp.strStartTime, lp.datEndDate, lp.strEndTime, lp.datConfirmed,
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
/****** Object:  StoredProcedure [dbo].[pr_evc_hr_report_base_data]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[pr_evc_hr_report_base_data]
AS
	SET NOCOUNT ON
	SELECT wp.WWID, wp.EmployeeStatusCd, CASE WHEN wp.FLSACd = 'E' THEN 'E' ELSE 'N' END AS FLSACd, 
	       wp.LastNm + ', ' + wp.FirstNm AS FullName, wp.NextLevelNm, wp.WorkLocationSiteCd, wp.OriginalStartDt,
  	       wp.StartDt, et.blnBlueBadge, u.datDOB,u.endDate,wp.DepartmentNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						ON wp.WWID = u.strWWID 
					       INNER JOIN dbo.tblEeType et 
					        ON wp.EmpTypeCode = et.strEmpTypeCode
	 WHERE et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
        ORDER BY FullName ASC
	RETURN


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_hr_report_employee_data]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[pr_evc_hr_report_employee_data](@vWWID	AS CHAR(8))
AS
	SET NOCOUNT ON
	SELECT wp.WWID, wp.EmployeeStatusCd, CASE WHEN wp.FLSACd = 'E' THEN 'E' ELSE 'N' END AS FLSACd, 
	       wp.LastNm + ', ' + wp.FirstNm AS FullName, wp.NextLevelNm, wp.WorkLocationSiteCd, wp.OriginalStartDt,
  	       wp.StartDt, et.blnBlueBadge, u.datDOB,u.endDate,wp.DepartmentNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						ON wp.WWID = u.strWWID 
					       INNER JOIN dbo.tblEeType et 
					        ON wp.EmpTypeCode = et.strEmpTypeCode
	 WHERE wp.WWID = @vWWID
		
	RETURN


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_hr_report_manager_reports]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[pr_evc_hr_report_manager_reports](@vWWID	AS CHAR(8))
AS
	SET NOCOUNT ON
	SELECT wp.WWID, wp.EmployeeStatusCd, CASE WHEN wp.FLSACd = 'E' THEN 'E' ELSE 'N' END AS FLSACd, 
	       wp.LastNm + ', ' + wp.FirstNm AS FullName, wp.NextLevelNm, wp.WorkLocationSiteCd, wp.OriginalStartDt,
  	       wp.StartDt, et.blnBlueBadge, u.datDOB,u.endDate,wp.DepartmentNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						ON wp.WWID = u.strWWID 
					       INNER JOIN dbo.tblEeType et 
					        ON wp.EmpTypeCode = et.strEmpTypeCode
	 WHERE wp.NextLevelWWID = @vWWID
		
	RETURN


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_payroll_report_filtered_site_only]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[pr_evc_payroll_report_filtered_site_only](@vStartDate 	DATETIME,
					                      @vEndDate		DATETIME,
					                      @vSite	        CHAR(2))
AS
	SET NOCOUNT ON
	SELECT wp.WWID, wp.LastNm + ', ' + wp.FirstNm AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.DepartmentNm, wp.NextLevelNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.CompanyCd = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
         AND l.datStartDate <= @vEndDate
  	   AND wp.WorkLocationSiteCd = @vSite
         ORDER BY wp.WWID ASC, Startdate ASC
	RETURN
  				


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_payroll_report_filtered_with_site]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[pr_evc_payroll_report_filtered_with_site](@vStartDate 	DATETIME,
					                      @vEndDate		DATETIME,
					                      @vStatus	        VARCHAR(20),
					                      @vSite	        CHAR(2))
AS
	SET NOCOUNT ON
	IF @vStatus = 'active'
	BEGIN
	SELECT wp.WWID, wp.LastNm + ', ' + wp.FirstNm AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.DepartmentNm, wp.NextLevelNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.CompanyCd = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
           AND l.datStartDate <= @vEndDate
	   AND wp.EmployeeStatusCd NOT IN ('R', 'T')
  	   AND wp.WorkLocationSiteCd = @vSite
         ORDER BY wp.WWID ASC, Startdate ASC
	END
	ELSE
	BEGIN
	SELECT wp.WWID, wp.LastNm + ', ' + wp.FirstNm AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.DepartmentNm, wp.NextLevelNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.CompanyCd = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
           AND l.datStartDate <= @vEndDate
	   AND wp.EmployeeStatusCd IN ('R', 'T')
  	   AND wp.WorkLocationSiteCd = @vSite
         ORDER BY wp.WWID ASC, Startdate ASC
        END
	RETURN
 


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_payroll_report_filtered_without_site]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[pr_evc_payroll_report_filtered_without_site](@vStartDate 	DATETIME,
					                         @vEndDate	DATETIME,
					                         @vStatus	VARCHAR(20))
AS
	SET NOCOUNT ON
	IF @vStatus = 'active'
	BEGIN
	SELECT wp.WWID, wp.LastNm + ', ' + wp.FirstNm AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.DepartmentNm, wp.NextLevelNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.CompanyCd = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
         AND l.datStartDate <= @vEndDate
	   AND wp.EmployeeStatusCd NOT IN ('R', 'T')
         ORDER BY wp.WWID ASC, Startdate ASC
	END
	ELSE
	BEGIN
	SELECT wp.WWID, wp.LastNm + ', ' + wp.FirstNm AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected,wp.DepartmentNm, wp.NextLevelNm
	  FROM dbo.qryWorkerPrivateAll wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.CompanyCd = 508
	   AND et.blnBlueBadge = 1
	   AND u.blnIsEELeaveTracked = 1
           AND ((l.datStartDate BETWEEN @vStartDate AND @vEndDate)
	    OR (l.datEndDate BETWEEN @vStartDate AND @vEndDate)
            OR (l.datStartDate < @vStartDate AND l.datEndDate > @vEndDate))
      AND l.datStartDate <= @vEndDate
	   AND wp.EmployeeStatusCd IN ('R', 'T')
         ORDER BY wp.WWID ASC, Startdate ASC
        END
	RETURN
  				


GO
/****** Object:  StoredProcedure [dbo].[pr_evc_shannon_sites]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/* add by MFILLAST */
CREATE PROCEDURE [dbo].[pr_evc_shannon_sites]
AS
	SET NOCOUNT ON
	SELECT DISTINCT wp.WorkLocationSiteCd 
	  FROM tblUser u INNER JOIN dbo.qryWorkerPrivateAll wp
	                         ON u.strWWID = wp.WWID
	 WHERE wp.CompanyCd = 508
	ORDER BY wp.WorkLocationSiteCd ASC
	RETURN


GO
/****** Object:  StoredProcedure [dbo].[pub_hol_add]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[pub_hol_add](@vMyDate DATETIME,@vName AS VARCHAR(50))
AS

	SET NOCOUNT ON
INSERT INTO tblPublicHoliday 
VALUES (@vMyDate,@vName)


GO
/****** Object:  StoredProcedure [dbo].[pub_hol_display]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[pub_hol_display]
AS

	SET NOCOUNT ON

Select *
From tblPublicHoliday


GO
/****** Object:  StoredProcedure [dbo].[public_hol_delete]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[public_hol_delete](@vMyDate DATETIME)
as

set nocount on

delete from tblPublicHoliday where datDate = @vMyDate


GO
/****** Object:  StoredProcedure [dbo].[public_hol_display]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[public_hol_display]
as

set nocount on

select datDate, strDescription from tblPublicHoliday


GO
/****** Object:  StoredProcedure [dbo].[team_emps_emps]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
Create procedure [dbo].[team_emps_emps](@vWWID AS Char(8))
as

set nocount on

select distinct WWID, FirstNm, LastNm from WorkerPrivate
Where NextLevelWWID = @vWWID AND EmployeeStatusCd = 'A'


GO
/****** Object:  StoredProcedure [dbo].[update_employees]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
create procedure [dbo].[update_employees](@vWWID As Char(8),@vEmployeeStatusCd As Char(1),@vNextLevelNm As VARCHAR,@vNextLevelWWID As Char(8))
As

set nocount on

update WorkerPrivate
set NextLevelNm = @vNextLevelNm, NextLevelWWID = @vNextLevelWWID, EmployeeStatusCd = @vEmployeeStatusCd

where WWID = @vWWID


GO
/****** Object:  StoredProcedure [dbo].[usp_approve_cancel_request]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 18/04/2002 14:09:03 ******/
/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 25/02/2002 16:35:40 ******/
/****** Object:  Stored Procedure dbo.usp_approve_cancel_request    Script Date: 01/06/2001 19:10:44 ******/
CREATE PROCEDURE [dbo].[usp_approve_cancel_request]
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
/****** Object:  StoredProcedure [dbo].[usp_approve_leave_request]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 18/04/2002 14:09:03 ******/
/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 25/02/2002 16:35:40 ******/
/****** Object:  Stored Procedure dbo.usp_approve_leave_request    Script Date: 01/06/2001 19:10:39 ******/
CREATE PROCEDURE [dbo].[usp_approve_leave_request]
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
/****** Object:  StoredProcedure [dbo].[usp_cancel_leave_request]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 18/04/2002 14:09:03 ******/
/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 25/02/2002 16:35:40 ******/
/****** Object:  Stored Procedure dbo.usp_cancel_leave_request    Script Date: 01/06/2001 19:10:44 ******/
CREATE PROCEDURE [dbo].[usp_cancel_leave_request]
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
/****** Object:  StoredProcedure [dbo].[usp_carryover_by_ID]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 18/04/2002 14:09:03 ******/
/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 25/02/2002 16:35:40 ******/
/****** Object:  Stored Procedure dbo.usp_carryover_by_ID    Script Date: 01/06/2001 19:10:39 ******/
CREATE PROCEDURE [dbo].[usp_carryover_by_ID]
@lngID integer
AS
SELECT * FROM tblcarryover WHERE lngID = @lngID


GO
/****** Object:  StoredProcedure [dbo].[usp_carryover_by_WWID_Year]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 18/04/2002 14:09:03 ******/
/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 25/02/2002 16:35:40 ******/
/****** Object:  Stored Procedure dbo.usp_carryover_by_WWID_Year    Script Date: 01/06/2001 19:10:39 ******/
CREATE PROCEDURE [dbo].[usp_carryover_by_WWID_Year]
@strEEWWID char(8),
@lngYear integer
AS
SELECT * FROM tblcarryovereoy WHERE strEEWWID = @strEEWWID and lngYear = @lngYear


GO
/****** Object:  StoredProcedure [dbo].[usp_carryovers_by_WWID]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 18/04/2002 14:09:03 ******/
/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 25/02/2002 16:35:41 ******/
/****** Object:  Stored Procedure dbo.usp_carryovers_by_WWID    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE [dbo].[usp_carryovers_by_WWID]
@strEEWWID char(8)
AS
SELECT * FROM tblcarryover WHERE strEEWWID = @strEEWWID
ORDER BY lngYear

GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE PROCEDURE [dbo].[usp_share_with_team_calendar_on]
@strEEWWID char(8)
AS
Update tblUser 
set shareWithTeamCalendar = 1
where strWWID = @strEEWWID
GO

drop procedure [usp_share_with_team_calendar_off]
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE PROCEDURE [dbo].[usp_share_with_team_calendar_off]
@strEEWWID char(8)
AS
Update tblUser 
set tblUser.shareWithTeamCalendar = 0
from tblUser tu, tblLeavePeriod tlp
where tu.strWWID = tlp.strEEWWID
and tu.strWWID = @strEEWWID;

Update tblLeavePeriod 
set tblLeavePeriod.shareLeaveWithTeamCalendar = 0
from tblUser tu, tblLeavePeriod tlp
where tu.strWWID = tlp.strEEWWID
and tlp.strEEWWID = @strEEWWID;
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE PROCEDURE [dbo].[usp_get_share_with_team_calendar]
@strEEWWID char(8)
AS
select shareWithTeamCalendar from tblUser where strWWID = @strEEWWID
return

GO

/****** Object:  StoredProcedure [dbo].[usp_create_user]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 18/04/2002 14:09:04 ******/
/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 25/02/2002 16:35:41 ******/
/****** Object:  Stored Procedure dbo.usp_create_user    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE [dbo].[usp_create_user]
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
/****** Object:  StoredProcedure [dbo].[usp_delete_leave_request]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 02/05/2002 14:23:59 ******/
/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 18/04/2002 14:09:04 ******/
/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 25/02/2002 16:35:41 ******/
/****** Object:  Stored Procedure dbo.usp_delete_leave_request    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE [dbo].[usp_delete_leave_request]
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

CREATE PROCEDURE [dbo].[usp_ee_payrollreportlite2]
(
	@datRptStartDate datetime,
	@datRptEndDate	datetime
)
AS
	SELECT wp.WWID, wp.LastNm + ', ' + wp.FirstNm AS FullName, lv.strLeaveTypeName as LeaveType,
	       l.datStartDate as StartDate, l.strStartTime as StartTime, l.datEndDate as EndDate,
	       l.strEndTime as EndTime, l.datRaised as LeaveRequestRaised, l.datApproved as Approved,
  	       l.datRejected as Rejected, l.datCancelRequested as CancelRequested, 
	       l.datCancelApproved as CancelApproved, l.datCancelRejected as CancelRejected
	  FROM WorkerPrivate wp INNER JOIN dbo.tblUser u 
						       ON wp.WWID = u.strWWID 
				               INNER JOIN dbo.tblEeType et 
					               ON wp.EmpTypeCode = et.strEmpTypeCode
                                               INNER JOIN tblLeavePeriod l 
			                               ON wp.WWID = l.strEEWWID 
			                       INNER JOIN tblLeaveType lv 
				                       ON l.lngLeaveTypeId = lv.lngID
			
	 WHERE wp.CompanyCd = 508
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
/****** Object:  StoredProcedure [dbo].[usp_eeapprovalspending]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 02/05/2002 14:24:00 ******/
/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 18/04/2002 14:09:04 ******/
/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eeapprovalspending    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE [dbo].[usp_eeapprovalspending]
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
/****** Object:  StoredProcedure [dbo].[usp_eedelegateformanagers]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 02/05/2002 14:24:00 ******/
/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 18/04/2002 14:09:04 ******/
/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eedelegateformanagers    Script Date: 01/06/2001 19:10:45 ******/
CREATE PROCEDURE [dbo].[usp_eedelegateformanagers]
@strWWID char(8)
AS
SELECT *
FROM qryEEDetails
WHERE strDelegateWWID = @strWWID


GO
/****** Object:  StoredProcedure [dbo].[usp_eedetails]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 02/05/2002 14:24:00 ******/
/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eedetails    Script Date: 01/06/2001 19:10:45 ******/
CREATE PROCEDURE [dbo].[usp_eedetails]
@strWWID char(8)
AS
SELECT *
FROM qryEEDetails
WHERE WWID = @strWWID


GO
/****** Object:  StoredProcedure [dbo].[usp_eedirectreports]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 02/05/2002 14:24:00 ******/
/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eedirectreports    Script Date: 01/06/2001 19:10:45 ******/
CREATE PROCEDURE [dbo].[usp_eedirectreports]
@strWWID char(8)
AS
SELECT *
FROM qryEEDetails 
WHERE NextLevelWWID = @strWWID


GO
/****** Object:  StoredProcedure [dbo].[usp_eeelp]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eeelp    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE [dbo].[usp_eeelp]
@strWWID char(8)
AS
SELECT *
FROM qryELPDetails
WHERE strEEWWID = @strWWID


GO
/****** Object:  StoredProcedure [dbo].[usp_eeleaveperiods_by_type_for_year]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE [dbo].[usp_eeleaveperiods_by_type_for_year]
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
/****** Object:  StoredProcedure [dbo].[usp_eeleaveperiods_by_type_for_year2]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaveperiods_by_type_for_year    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE [dbo].[usp_eeleaveperiods_by_type_for_year2]
@strWWID char(8),
@strLeaveTypeName char(30),
@newDate smalldatetime,
@oldDate smalldatetime,
@strYear char(8)
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
/****** Object:  StoredProcedure [dbo].[usp_eeleaverequests]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE [dbo].[usp_eeleaverequests]
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
/****** Object:  StoredProcedure [dbo].[usp_eeleaverequests_adminview]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 25/02/2002 16:35:42 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_adminview    Script Date: 01/06/2001 19:10:49 ******/
CREATE PROCEDURE [dbo].[usp_eeleaverequests_adminview]
@strWWID char(8)
AS
Select * From qryLeavePeriodDetails
WHERE strEEWWID = @strWWID
ORDER BY datStartDate, strStartTime, datEndDate, strEndTime


GO
/****** Object:  StoredProcedure [dbo].[usp_eeleaverequests_overlapping]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 25/02/2002 16:35:43 ******/
/****** Object:  Stored Procedure dbo.usp_eeleaverequests_overlapping    Script Date: 01/06/2001 19:10:50 ******/
CREATE PROCEDURE [dbo].[usp_eeleaverequests_overlapping]
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
/****** Object:  StoredProcedure [dbo].[usp_eesearchbyname]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 25/02/2002 16:35:43 ******/
/****** Object:  Stored Procedure dbo.usp_eesearchbyname    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE [dbo].[usp_eesearchbyname]
@strName varchar(50),
@lngMaxResults integer
AS
Set @strName = RTrim(@strName)
Set @lngMaxResults = @lngMaxResults + 1
SET ROWCOUNT @lngMaxResults
SELECT WWID, LastNm, FirstNm, NextLevelNm, MailStopTxt FROM qryEEDetails
WHERE LastNm+","+FirstNm Like(@strName + '%')
ORDER BY LastNm, FirstNm
Return @@rowcount


GO
/****** Object:  StoredProcedure [dbo].[usp_eeshowall]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[usp_eeshowall]
AS
SELECT WWID, LastNm, FirstNm, NextLevelNm, MailStopTxt, EndDt FROM qryEEDetails
WHERE not EndDt < SYSDATETIME()
OR EndDt is NULL
OR EndDt = '9999-12-31 00:00:00.000'
ORDER BY LastNm, FirstNm
Return @@rowcount



GO
/****** Object:  StoredProcedure [dbo].[usp_elpinstance]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 02/05/2002 14:24:01 ******/
/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 18/04/2002 14:09:05 ******/
/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 25/02/2002 16:35:43 ******/
/****** Object:  Stored Procedure dbo.usp_elpinstance    Script Date: 01/06/2001 19:10:50 ******/
CREATE PROCEDURE [dbo].[usp_elpinstance]
@lngID integer
AS
SELECT *
FROM qryELPDetails
WHERE ELPID = @lngID


GO
/****** Object:  StoredProcedure [dbo].[usp_elpinstance_elpreliefs]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 25/02/2002 16:35:43 ******/
/****** Object:  Stored Procedure dbo.usp_elpinstance_elpreliefs    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE [dbo].[usp_elpinstance_elpreliefs]
@lngELPID integer
AS
SELECT *
FROM dbo.tblELPRelief 
WHERE lngELPID = @lngELPID
ORDER BY lngYear


GO
/****** Object:  StoredProcedure [dbo].[usp_elprelief]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 25/02/2002 16:35:43 ******/
/****** Object:  Stored Procedure dbo.usp_elprelief    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE [dbo].[usp_elprelief]
@lngID integer
AS
	SELECT *
	FROM dbo.tblELPRelief 
	WHERE lngID = @lngID


GO
/****** Object:  StoredProcedure [dbo].[usp_getLastUpdateWdsCopy_from_shortid]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/* added by MFILLAST : return the date of the last update of info stored in the copy of wds(using user's shortid). */
CREATE PROCEDURE [dbo].[usp_getLastUpdateWdsCopy_from_shortid]
@strIdsid char(8)
AS
SET NOCOUNT ON
SELECT WdsCopyUpdateDate
FROM WorkerPrivate
WHERE Idsid = @strIdsid


GO
/****** Object:  StoredProcedure [dbo].[usp_getLastUpdateWdsCopy_from_wwid]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/* added by MFILLAST : return the date of the last update of info stored in the copy of wds(using user's wwid). */
CREATE PROCEDURE [dbo].[usp_getLastUpdateWdsCopy_from_wwid]
@strWWID char(8)
AS
SET NOCOUNT ON
SELECT WdsCopyUpdateDate
FROM WorkerPrivate
WHERE WWID = @strWWID


GO
/****** Object:  StoredProcedure [dbo].[usp_insert_wds_copy]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[usp_insert_wds_copy]
@WWID char(8)  ,
@FirstNm varchar(30) ,
@LastNm varchar(30) , 
@Idsid char(8) ,
@EmployeeStatusCd char(1) ,
@CorporateEmailTxt varchar(80) ,
@CompanyCd char(3) ,
@MailStopTxt varchar(50) ,
@NextLevelNm varchar(50) ,
@NextLevelWWID char(8) ,
@RegionCode varchar(4) ,
@WorkLocationSiteCd char(2) ,
@EndDt datetime ,
@StartDt smalldatetime ,
@EmpTypeCode char(3) ,
@OriginalStartDt smalldatetime  ,
@FullTmPartTmCd varchar(2) ,
@FLSACd varchar(1),
@DepartmentNm varchar(50) AS
SET NOCOUNT ON
INSERT INTO WorkerPrivate
( WWID ,FirstNm ,LastNm , Idsid ,EmployeeStatusCd ,CorporateEmailTxt ,CompanyCd ,MailStopTxt ,NextLevelNm ,NextLevelWWID ,RegionCode ,WorkLocationSiteCd ,EndDt ,StartDt ,EmpTypeCode ,OriginalStartDt ,FullTmPartTmCd ,FLSACd,DepartmentNm,WdsCopyUpdateDate )
VALUES (@WWID ,@FirstNm ,@LastNm ,@Idsid ,@EmployeeStatusCd ,@CorporateEmailTxt ,@CompanyCd ,@MailStopTxt ,@NextLevelNm ,@NextLevelWWID ,@RegionCode ,@WorkLocationSiteCd ,@EndDt ,@StartDt ,@EmpTypeCode ,@OriginalStartDt ,@FullTmPartTmCd ,@FLSACd,@DepartmentNm,GetDate() )
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred inserting wds copy.'
	return(0)
end
else
begin
	   -- Return 1 to the calling program to indicate success.
	print 'inserted in local copy of wds .'
	return(1)
end


GO
/****** Object:  StoredProcedure [dbo].[usp_leaveperiod]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 25/02/2002 16:35:43 ******/
/****** Object:  Stored Procedure dbo.usp_leaveperiod    Script Date: 01/06/2001 19:10:40 ******/
CREATE PROCEDURE [dbo].[usp_leaveperiod]
@lngID integer
AS
SELECT *
FROM dbo.tblLeavePeriod WHERE lngID = @lngID


GO
/****** Object:  StoredProcedure [dbo].[usp_leavetype_by_id]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 25/02/2002 16:35:43 ******/
/****** Object:  Stored Procedure dbo.usp_leavetype_by_id    Script Date: 01/06/2001 19:10:41 ******/
CREATE PROCEDURE [dbo].[usp_leavetype_by_id]
@lngID integer
AS
	SELECT *
	FROM dbo.tblLeaveType
	WHERE lngID = @lngID


GO
/****** Object:  StoredProcedure [dbo].[usp_leavetype_by_name]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_leavetype_by_name    Script Date: 01/06/2001 19:10:41 ******/
CREATE PROCEDURE [dbo].[usp_leavetype_by_name]
@strName char(30)
AS
	SELECT *
	FROM dbo.tblLeaveType
	WHERE strLeaveTypeName = @strName


GO
/****** Object:  StoredProcedure [dbo].[usp_leavetypes_admin]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_admin    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE [dbo].[usp_leavetypes_admin] AS
SELECT dbo.tblLeaveType.*
FROM dbo.tblLeaveType
ORDER BY strLeaveTypeName


GO
/****** Object:  StoredProcedure [dbo].[usp_leavetypes_ee_other_leave]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_other_leave    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE [dbo].[usp_leavetypes_ee_other_leave] AS
SELECT dbo.tblLeaveType.*
FROM dbo.tblLeaveType
WHERE blnIsOtherLeave = 1
ORDER BY strLeaveTypeName


GO
/****** Object:  StoredProcedure [dbo].[usp_leavetypes_ee_requests]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_leavetypes_ee_requests    Script Date: 01/06/2001 19:10:42 ******/
CREATE PROCEDURE [dbo].[usp_leavetypes_ee_requests] AS
SELECT dbo.tblLeaveType.*
FROM dbo.tblLeaveType
WHERE blnEERequests = 1


GO
/****** Object:  StoredProcedure [dbo].[usp_publicholidays]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 18/04/2002 14:09:06 ******/
/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_publicholidays    Script Date: 01/06/2001 19:10:43 ******/
CREATE PROCEDURE [dbo].[usp_publicholidays]
AS
Select * From tblPublicHoliday
ORDER BY datDate


GO
/****** Object:  StoredProcedure [dbo].[usp_reject_cancel_request]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 02/05/2002 14:24:02 ******/
/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_reject_cancel_request    Script Date: 01/06/2001 19:10:46 ******/
CREATE PROCEDURE [dbo].[usp_reject_cancel_request]
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
/****** Object:  StoredProcedure [dbo].[usp_reject_leave_request]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_reject_leave_request    Script Date: 01/06/2001 19:10:41 ******/
CREATE PROCEDURE [dbo].[usp_reject_leave_request]
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
/****** Object:  StoredProcedure [dbo].[usp_save_carryover]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 25/02/2002 16:35:44 ******/
/****** Object:  Stored Procedure dbo.usp_save_carryover    Script Date: 01/06/2001 19:10:43 ******/
CREATE PROCEDURE [dbo].[usp_save_carryover]
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
/****** Object:  StoredProcedure [dbo].[usp_save_elp_activation]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 25/02/2002 16:35:45 ******/
/****** Object:  Stored Procedure dbo.usp_save_elp_activation    Script Date: 01/06/2001 19:10:43 ******/
CREATE PROCEDURE [dbo].[usp_save_elp_activation]
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
/****** Object:  StoredProcedure [dbo].[usp_save_elprelief]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 25/02/2002 16:35:45 ******/
/****** Object:  Stored Procedure dbo.usp_save_elprelief    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE [dbo].[usp_save_elprelief]
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
/****** Object:  StoredProcedure [dbo].[usp_save_leave_request_admin]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 25/02/2002 16:35:45 ******/
/****** Object:  Stored Procedure dbo.usp_save_leave_request_admin    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE [dbo].[usp_save_leave_request_admin]
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
@lngELPID integer,
@lngCompTimeID integer,
@shareLeaveWithTeamCalendar bit
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
			lngELPID,
			lngCompTimeID,
			shareLeaveWithTeamCalendar
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
			@lngELPID,
			@lngCompTimeID,
			@shareLeaveWithTeamCalendar
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
			lngELPID = @lngELPID,
			lngCompTimeID = @lngCompTimeID,
			shareLeaveWithTeamCalendar = @shareLeaveWithTeamCalendar
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
/****** Object:  StoredProcedure [dbo].[usp_save_new_comp_time]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 25/02/2002 16:35:45 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE [dbo].[usp_save_new_comp_time]
@lngWWID integer,
@lngDaysGranted integer,
@strReason varchar(100)
AS
insert into tblCompTime
	(
	lngWWID,
	lngDaysGranted,
 	datDateGranted,
	strReason
	)
Values (
	@lngWWID,
	@lngDaysGranted,
	getDate(),
	@strReason
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
/****** Object:  StoredProcedure [dbo].[usp_save_new_leave_request_standard]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 25/02/2002 16:35:45 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE [dbo].[usp_save_new_leave_request_standard]
@strEEWWID char(8),
@strApproverWWID char(8),
@lngLeaveTypeID integer,
@datStartDate smalldatetime,
@strStartTime char(2),
@datEndDate smalldatetime,
@strEndTime char(2),
@strRequestComments varchar(100),
@lngELPID integer,
@lngCompTimeID integer,
@shareLeaveWithTeamCalendar bit
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
	lngELPID,
	lngCompTimeID,
	shareLeaveWithTeamCalendar
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
	@lngELPID,
	@lngCompTimeID,
	@shareLeaveWithTeamCalendar
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
/****** Object:  StoredProcedure [dbo].[usp_save_new_leave_request_standard_Comp]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 25/02/2002 16:35:45 ******/
/****** Object:  Stored Procedure dbo.usp_save_new_leave_request_standard    Script Date: 01/06/2001 19:10:47 ******/
CREATE PROCEDURE [dbo].[usp_save_new_leave_request_standard_Comp]
@strEEWWID char(8),
@strApproverWWID char(8),
@lngLeaveTypeID integer,
@datStartDate smalldatetime,
@strStartTime char(2),
@datEndDate smalldatetime,
@strEndTime char(2),
@strRequestComments varchar(100),
@lngCompTimeID integer,
@shareLeaveWithTeamCalendar bit
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
	lngCompTimeID,
	shareLeaveWithTeamCalendar
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
	@lngCompTimeID,
	@shareLeaveWithTeamCalendar
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
/****** Object:  StoredProcedure [dbo].[usp_save_user]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[usp_save_user]
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

/****** Object:  StoredProcedure [dbo].[usp_update_wds_copy]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[usp_update_wds_copy]
@WWID char(8)  ,
@FirstNm varchar(30) ,
@LastNm varchar(30) , 
@Idsid char(8) ,
@EmployeeStatusCd char(1) ,
@CorporateEmailTxt varchar(80) ,
@CompanyCd char(3) ,
@MailStopTxt varchar(50) ,
@NextLevelNm varchar(50) ,
@NextLevelWWID char(8) ,
@RegionCode varchar(4) ,
@WorkLocationSiteCd char(2) ,
@EndDt datetime ,
@StartDt smalldatetime ,
@EmpTypeCode char(3) ,
@OriginalStartDt smalldatetime  ,
@FullTmPartTmCd varchar(2) ,
@FLSACd varchar(1),
@DepartmentNm varchar(50) AS
SET NOCOUNT ON
UPDATE WorkerPrivate
SET 
Idsid = @Idsid  ,
FirstNm = @FirstNm  ,
LastNm = @LastNm  , 
EmployeeStatusCd = @EmployeeStatusCd  ,
CorporateEmailTxt = @CorporateEmailTxt  ,
CompanyCd = @CompanyCd  ,
MailStopTxt = @MailStopTxt  ,
NextLevelNm = @NextLevelNm  ,
NextLevelWWID = @NextLevelWWID  ,
RegionCode = @RegionCode  ,
WorkLocationSiteCd = @WorkLocationSiteCd  ,
EndDt = @EndDt  ,
StartDt = @StartDt  ,
EmpTypeCode = @EmpTypeCode  ,
OriginalStartDt = @OriginalStartDt   ,
FullTmPartTmCd = @FullTmPartTmCd  ,
FLSACd = @FLSACd  ,
DepartmentNm =@DepartmentNm,
WdsCopyUpdateDate = GetDate()
WHERE WWID = @WWID
Declare @lngError int
Select @lngError = @@ERROR
If @lngError <> 0
begin
	-- Return the error code to the calling program to indicate failure.
	print 'An error occurred updating wds copy.'
	return(0)
end
else
begin
	   -- Return 1 to the calling program to indicate success.
	print 'Local copy of wds updated.'
	return(1)
end


GO
/****** Object:  StoredProcedure [dbo].[usp_userdetails]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 02/05/2002 14:24:03 ******/
/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 18/04/2002 14:09:07 ******/
/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 25/02/2002 16:35:45 ******/
/****** Object:  Stored Procedure dbo.usp_userdetails    Script Date: 01/06/2001 19:10:48 ******/
CREATE PROCEDURE [dbo].[usp_userdetails]
@strIdsid char(8)
AS
SELECT *
FROM qryEEDetails
WHERE Idsid = @strIdsid


GO
/****** Object:  Table [dbo].[development]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[development](
	[username] [varchar](50) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[fkeys]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[fkeys](
	[PKTABLE_NAME] [varchar](32) NULL,
	[FKTABLE_NAME] [varchar](32) NULL,
	[FKCOLUMN1_NAME] [varchar](32) NULL,
	[FKCOLUMN2_NAME] [varchar](32) NULL,
	[FKCOLUMN3_NAME] [varchar](32) NULL,
	[FKCOLUMN4_NAME] [varchar](32) NULL,
	[FKCOLUMN5_NAME] [varchar](32) NULL,
	[FKCOLUMN6_NAME] [varchar](32) NULL,
	[FKCOLUMN7_NAME] [varchar](32) NULL,
	[FKCOLUMN8_NAME] [varchar](32) NULL,
	[FKCOLUMN9_NAME] [varchar](32) NULL,
	[FKCOLUMN10_NAME] [varchar](32) NULL,
	[FKCOLUMN11_NAME] [varchar](32) NULL,
	[FKCOLUMN12_NAME] [varchar](32) NULL,
	[FKCOLUMN13_NAME] [varchar](32) NULL,
	[FKCOLUMN14_NAME] [varchar](32) NULL,
	[FKCOLUMN15_NAME] [varchar](32) NULL,
	[FKCOLUMN16_NAME] [varchar](32) NULL,
	[FK_NAME] [varchar](32) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblCarryOver]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblCarryOver](
	[lngID] [int] IDENTITY(1,1) NOT NULL,
	[strEEWWID] [char](8) NOT NULL,
	[lngYear] [smallint] NOT NULL,
	[lngDays] [real] NULL CONSTRAINT [DF_tblCarryOver_lngDays]  DEFAULT ((0)),
	[strEnteredByWWID] [char](8) NULL CONSTRAINT [DF_tblCarryOverEOY_strEnteredByWWID]  DEFAULT ('System'),
	[datEntered] [datetime] NULL CONSTRAINT [DF_tblCarryOverEOY_datEntered]  DEFAULT (getdate()),
	[strComments] [varchar](100) NULL,
	[blnPreArranged] [bit] NULL CONSTRAINT [DF_tblCarryOver_blnPreArranged]  DEFAULT ((0)),
 CONSTRAINT [PK_tblCarryOver] PRIMARY KEY CLUSTERED 
(
	[lngID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblCompTime]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblCompTime](
	[lngID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[lngWWID] [int] NOT NULL,
	[lngDaysGranted] [int] NULL CONSTRAINT [DF_tblCompTime_lngDaysGranted]  DEFAULT ((0)),
	[datDateGranted] [datetime] NULL,
	[datDateRevoked] [datetime] NULL,
	[IngDaysRevoked] [int] NULL CONSTRAINT [DF_tblCompTime_IngDaysRevoked]  DEFAULT ((0)),
	[lngDaysBooked] [int] NULL CONSTRAINT [DF_tblCompTime_lngDaysBooked]  DEFAULT ((0)),
	[Taken] [int] NULL CONSTRAINT [DF_tblCompTime_Taken]  DEFAULT ((0)),
	[strReason] [varchar](500) NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblEeType]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblEeType](
	[lngID] [int] IDENTITY(1,1) NOT NULL,
	[strEmpTypeCode] [char](3) NOT NULL,
	[strDescription] [varchar](30) NOT NULL,
	[blnBlueBadge] [bit] NOT NULL,
 CONSTRAINT [PK_tblEeType] PRIMARY KEY NONCLUSTERED 
(
	[lngID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblELPInstance]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblELPInstance](
	[lngID] [int] IDENTITY(1,1) NOT NULL,
	[strEEWWID] [char](8) NULL,
	[datActivated] [datetime] NULL,
	[strActivatedByWWID] [char](8) NULL,
	[dblInitialDaysBanked] [real] NULL,
	[dblTargetDays] [real] NULL,
 CONSTRAINT [PK_tblELPInstance] PRIMARY KEY NONCLUSTERED 
(
	[lngID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblELPRelief]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblELPRelief](
	[lngID] [int] IDENTITY(1,1) NOT NULL,
	[lngELPID] [int] NULL,
	[lngYear] [smallint] NULL,
	[datEntered] [datetime] NULL,
	[strEnteredByWWID] [char](8) NULL,
	[strComments] [varchar](50) NULL,
 CONSTRAINT [PK_tblELPRelief] PRIMARY KEY NONCLUSTERED 
(
	[lngID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY],
 CONSTRAINT [IX_tblELPRelief_ELPID_Year] UNIQUE NONCLUSTERED 
(
	[lngELPID] ASC,
	[lngYear] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblLeavePeriod]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblLeavePeriod](
	[lngID] [int] IDENTITY(1,1) NOT NULL,
	[strEEWWID] [char](8) NULL,
	[strApproverWWID] [char](8) NULL,
	[lngLeaveTypeID] [int] NULL,
	[datStartDate] [smalldatetime] NULL,
	[strStartTime] [char](2) NULL,
	[datEndDate] [smalldatetime] NULL,
	[strEndTime] [char](2) NULL,
	[datRaised] [datetime] NOT NULL CONSTRAINT [DF_tblLeavePeriod_datRaised]  DEFAULT (getdate()),
	[datApproved] [datetime] NULL,
	[datRejected] [datetime] NULL,
	[datCancelRequested] [datetime] NULL,
	[datCancelApproved] [datetime] NULL,
	[datCancelRejected] [datetime] NULL,
	[datConfirmed] [datetime] NULL,
	[datConfirmEmailSent] [datetime] NULL,
	[strRequestComments] [varchar](100) NULL,
	[strResponseComments] [varchar](100) NULL,
	[lngELPID] [int] NULL,
	[lngCompTimeID] [int] NULL,
	[shareLeaveWithTeamCalendar] [bit] Default null,
 CONSTRAINT [PK_tblLeavePeriod] PRIMARY KEY NONCLUSTERED 
(
	[lngID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblLeaveType]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblLeaveType](
	[lngID] [int] IDENTITY(1,1) NOT NULL,
	[strLeaveTypeName] [varchar](30) NOT NULL,
	[blnEERequests] [bit] NULL CONSTRAINT [DF_tblLeaveType_blnEERequests]  DEFAULT ((0)),
	[blnAdminRequests] [bit] NULL CONSTRAINT [DF_tblLeaveType_blnAdminRequests]  DEFAULT ((0)),
	[blnRequestBeforeAccrued] [bit] NULL CONSTRAINT [DF_tblLeaveType_blnRequestBeforeAccrued]  DEFAULT ((0)),
	[dblMinimumDays] [real] NULL CONSTRAINT [DF_tblLeaveType_lngMinimumDays]  DEFAULT ((0)),
	[dblEntitlement] [real] NULL CONSTRAINT [DF_tblLeaveType_lngEntitlement]  DEFAULT ((0)),
	[dblDaysBeforeStopsLegalAdjAccrual] [real] NULL CONSTRAINT [DF_tblLeaveType_lngDaysBeforeStopsLegalAdjAccrual]  DEFAULT ((-1)),
	[blnDaysBeforeStopsIsConsecutive] [bit] NULL CONSTRAINT [DF_tblLeaveType_blnAllDaysIncludedBeforeStopsLegalAdjAccrual]  DEFAULT ((0)),
	[blnIsOtherLeave] [bit] NULL CONSTRAINT [DF_tblLeaveType_blnIsOtherLeave]  DEFAULT ((1)),
 CONSTRAINT [PK_tblLeaveType] PRIMARY KEY NONCLUSTERED 
(
	[lngID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY],
 CONSTRAINT [IX_tblLeaveType_strLeaveTypeName] UNIQUE NONCLUSTERED 
(
	[strLeaveTypeName] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblPublicHoliday]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblPublicHoliday](
	[lngID] [int] IDENTITY(1,1) NOT NULL,
	[datDate] [datetime] NULL,
	[strDescription] [varchar](50) NULL,
 CONSTRAINT [PK_tblPublicHoliday] PRIMARY KEY NONCLUSTERED 
(
	[lngID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[tblUser]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[tblUser](
	[strWWID] [char](8) NOT NULL,
	[datDOB] [datetime] NULL,
	[blnIsAdmin] [bit] NOT NULL CONSTRAINT [DF_tblUser_blnIsAdmin]  DEFAULT ((0)),
	[strDelegateWWID] [char](8) NULL,
	[blnIsException] [bit] NOT NULL CONSTRAINT [DF_tblUser_blnIsException]  DEFAULT ((0)),
	[strExceptionComments] [varchar](50) NULL,
	[datAquisitionDate] [smalldatetime] NULL,
	[blnIsExemptStatusChanged] [bit] NULL CONSTRAINT [DF_tblUser_blnExemptStatusChanged]  DEFAULT ((0)),
	[blnIsEELeaveTracked] [bit] NULL,
	[endDate] [datetime] NULL,
	[shareWithTeamCalendar] [bit] Default null,
 CONSTRAINT [PK_tblUser] PRIMARY KEY NONCLUSTERED 
(
	[strWWID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[WorkerPrivate]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[WorkerPrivate](
	[WWID] [char](8) NOT NULL,
	[FirstNm] [varchar](30) NULL,
	[LastNm] [varchar](30) NULL,
	[Idsid] [char](8) NULL,
	[EmployeeStatusCd] [char](1) NULL,
	[CorporateEmailTxt] [varchar](80) NULL,
	[CompanyCd] [char](3) NULL,
	[MailStopTxt] [varchar](50) NULL,
	[NextLevelNm] [varchar](50) NULL,
	[NextLevelWWID] [char](8) NULL,
	[RegionCode] [varchar](4) NULL,
	[WorkLocationSiteCd] [char](2) NULL,
	[EndDt] [datetime] NULL,
	[StartDt] [smalldatetime] NULL,
	[EmpTypeCode] [char](3) NOT NULL,
	[OriginalStartDt] [smalldatetime] NULL,
	[FullTmPartTmCd] [varchar](2) NULL,
	[FLSACd] [varchar](1) NULL,
	[WdsCopyUpdateDate] [smalldatetime] NULL,
	[DepartmentNm] [varchar](50) NULL,
 CONSTRAINT [PK_WorkerPrivate] PRIMARY KEY NONCLUSTERED 
(
	[WWID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  View [dbo].[qryWorkerPrivateAll]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[qryWorkerPrivateAll]
AS
SELECT *
FROM dbo.WorkerPrivate


GO
/****** Object:  View [dbo].[qryEEDetails]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[qryEEDetails]
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
/****** Object:  View [dbo].[qryPayrollReport]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  View dbo.qryPayrollReport    Script Date: 02/05/2002 14:23:57 ******/
/****** Object:  View dbo.qryPayrollReport    Script Date: 18/04/2002 14:09:01 ******/
CREATE VIEW [dbo].[qryPayrollReport] AS
SELECT 	u.WWID, 
	u.LastNm + ', ' + u.FirstNm   AS FullName, 
	lv.strLeaveTypeName as LeaveType, 
	l.datStartDate as StartDate, 
	l.strStartTime as StartTime, 
	l.datEndDate as EndDate, 
	l.strEndTime as EndTime,
	l.datCancelApproved as CancelApproved,
	u.EmployeeStatusCd as Status,
	u.FLSACd as ExemptStatus,
	u.StartDt as ServiceDate,
	et.blnBlueBadge as IsBlueBadge,
	u.OriginalStartDt as ODOH,
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
/****** Object:  View [dbo].[qryPayrollReport2]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  View dbo.qryPayrollReport2    Script Date: 02/05/2002 14:23:57 ******/
/****** Object:  View dbo.qryPayrollReport2    Script Date: 18/04/2002 14:09:01 ******/
CREATE VIEW [dbo].[qryPayrollReport2] AS
SELECT 	u.WWID, /* Payroll Rpt , HR Rpt, Admin Rpt */
	u.LastNm + ', ' + u.FirstNm   AS FullName,  /* Payroll Rpt, HR Rpt */
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
	u.EmployeeStatusCd as Status, /* Payroll Rpt, HR Rpt */
	u.EndDt as TeminationDate,
	u.FLSACd as ExemptStatus, /* HR Rpt, Admin Rpt */
	u.StartDt as LDOH, /* Admin Rpt */
	u.OriginalStartDt as ODOH, /* Admin Rpt */
	et.blnBlueBadge as IsBlueBadge,
	u.blnIsAdmin as IsAdmin, /* Admin Rpt */
	u.blnIsEELeaveTracked as IsEELeaveTracked,  /* Admin Rpt */
	
	ee.datDOB as DOB, /* Admin Rpt */
	u.NextLevelNm as ManagerName, /* HR Rpt */
	u.NextLevelWWID as ManagerWWID, /* HR Rpt */
	u.RegionCode as RegionCode, /* Payroll Rpt */
	u.WorkLocationSiteCd as WorkLocationSiteCd, /* Payroll Rpt */
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
/****** Object:  View [dbo].[qryUserApprovalsAll]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  View dbo.qryUserApprovalsAll    Script Date: 02/05/2002 14:23:57 *****
***** Object:  View dbo.qryUserApprovalsAll    Script Date: 18/04/2002 14:09:01 ******/
CREATE VIEW [dbo].[qryUserApprovalsAll]
AS
SELECT     EE.NextLevelWWID AS strEEManagerWWID, Manager.strDelegateWWID AS strManagerDelegateWWID, dbo.tblLeavePeriod.lngID, dbo.tblLeavePeriod.strEEWWID, 
                      dbo.tblLeavePeriod.strApproverWWID, dbo.tblLeavePeriod.lngLeaveTypeID, dbo.tblLeavePeriod.datStartDate, dbo.tblLeavePeriod.strStartTime, 
                      dbo.tblLeavePeriod.datEndDate, dbo.tblLeavePeriod.strEndTime, dbo.tblLeavePeriod.datRaised, dbo.tblLeavePeriod.datApproved, dbo.tblLeavePeriod.datRejected, 
                      dbo.tblLeavePeriod.datCancelRequested, dbo.tblLeavePeriod.datCancelApproved, dbo.tblLeavePeriod.datCancelRejected, dbo.tblLeavePeriod.datConfirmed, 
                      dbo.tblLeavePeriod.datConfirmEmailSent, dbo.tblLeavePeriod.strRequestComments, dbo.tblLeavePeriod.strResponseComments, dbo.tblLeavePeriod.lngELPID, 
                      dbo.tblLeavePeriod.lngCompTimeID
FROM         dbo.tblLeavePeriod LEFT OUTER JOIN
                      dbo.qryWorkerPrivateAll AS EE ON dbo.tblLeavePeriod.strEEWWID = EE.WWID LEFT OUTER JOIN
                      dbo.tblUser AS Manager ON EE.NextLevelWWID = Manager.strWWID
WHERE     (dbo.tblLeavePeriod.datStartDate > DATEADD(year, - 1, GETDATE()))


GO
/****** Object:  View [dbo].[qryActiveEmployee]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[qryActiveEmployee]
AS
SELECT DISTINCT 
                      TOP 100 PERCENT dbo.WorkerPrivate.WWID, dbo.WorkerPrivate.FirstNm, dbo.WorkerPrivate.LastNm, dbo.WorkerPrivate.NextLevelNm, 
                      dbo.WorkerPrivate.NextLevelWWID, dbo.WorkerPrivate.EmployeeStatusCd, dbo.WorkerPrivate.CompanyCd, dbo.WorkerPrivate.FLSACd, 
                      dbo.WorkerPrivate.StartDt, dbo.tblUser.endDate
FROM         dbo.WorkerPrivate INNER JOIN
                      dbo.tblUser ON dbo.WorkerPrivate.WWID = dbo.tblUser.strWWID
WHERE     (dbo.WorkerPrivate.FLSACd = 'E') AND (dbo.WorkerPrivate.CompanyCd = '508') AND (dbo.WorkerPrivate.EmployeeStatusCd = 'A')
ORDER BY dbo.WorkerPrivate.NextLevelNm, dbo.WorkerPrivate.WWID


GO
/****** Object:  View [dbo].[qryCarryOver]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[qryCarryOver]
AS
SELECT DISTINCT 
                      dbo.qryActiveEmployee.WWID, dbo.qryActiveEmployee.FirstNm, dbo.qryActiveEmployee.LastNm, dbo.tblCarryOver.lngYear, 
                      dbo.tblCarryOver.lngDays
FROM         dbo.tblCarryOver INNER JOIN
                      dbo.qryActiveEmployee ON dbo.tblCarryOver.strEEWWID = dbo.qryActiveEmployee.WWID
WHERE     (dbo.tblCarryOver.lngYear = 2009)


GO
/****** Object:  View [dbo].[qryEmployeeLeave]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[qryEmployeeLeave]
AS
SELECT DISTINCT 
                      dbo.qryActiveEmployee.WWID, dbo.qryActiveEmployee.FirstNm, dbo.qryActiveEmployee.LastNm, dbo.tblLeavePeriod.datStartDate, 
                      dbo.tblLeavePeriod.datEndDate, dbo.tblLeavePeriod.strStartTime, dbo.tblLeavePeriod.strEndTime
FROM         dbo.tblLeaveType INNER JOIN
                      dbo.tblLeavePeriod ON dbo.tblLeaveType.lngID = dbo.tblLeavePeriod.lngLeaveTypeID INNER JOIN
                      dbo.qryActiveEmployee ON dbo.tblLeavePeriod.strEEWWID = dbo.qryActiveEmployee.WWID
WHERE     (dbo.tblLeavePeriod.datStartDate > CONVERT(DATETIME, '2009-01-01 00:00:00', 102)) AND (dbo.tblLeavePeriod.datEndDate > CONVERT(DATETIME, 
                      '2009-01-01 00:00:00', 102))


GO
/****** Object:  View [dbo].[Admins]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[Admins]
AS
SELECT        dbo.WorkerPrivate.FirstNm, dbo.WorkerPrivate.LastNm, dbo.tblUser.blnIsAdmin
FROM            dbo.tblUser INNER JOIN
                         dbo.WorkerPrivate ON dbo.tblUser.strWWID = dbo.WorkerPrivate.WWID
WHERE        (dbo.tblUser.blnIsAdmin = 1)


GO
/****** Object:  View [dbo].[HRReport]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE VIEW [dbo].[HRReport]
AS
SELECT     dbo.WorkerPrivate.WWID, dbo.WorkerPrivate.FirstNm, dbo.WorkerPrivate.LastNm, dbo.WorkerPrivate.NextLevelNm, dbo.WorkerPrivate.NextLevelWWID, 
                      dbo.WorkerPrivate.StartDt, dbo.tblUser.endDate, dbo.tblCarryOver.lngYear, dbo.WorkerPrivate.EmployeeStatusCd, dbo.WorkerPrivate.CompanyCd, 
                      dbo.WorkerPrivate.FLSACd
FROM         dbo.WorkerPrivate INNER JOIN
                      dbo.tblUser ON dbo.WorkerPrivate.WWID = dbo.tblUser.strWWID INNER JOIN
                      dbo.tblCarryOver ON dbo.WorkerPrivate.WWID = dbo.tblCarryOver.strEEWWID INNER JOIN
                      dbo.tblLeavePeriod ON dbo.WorkerPrivate.WWID = dbo.tblLeavePeriod.strEEWWID INNER JOIN
                      dbo.tblELPInstance ON dbo.WorkerPrivate.WWID = dbo.tblELPInstance.strEEWWID
WHERE     (dbo.WorkerPrivate.EmployeeStatusCd = 'A') AND (dbo.WorkerPrivate.CompanyCd = '508') AND (dbo.WorkerPrivate.FLSACd = 'E') AND 
                      (dbo.tblCarryOver.lngYear = 2009)


GO
/****** Object:  View [dbo].[qryELPDetails]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
/****** Object:  View dbo.qryELPDetails    Script Date: 02/05/2002 14:23:57 ******/
/****** Object:  View dbo.qryELPDetails    Script Date: 18/04/2002 14:09:01 ******/
/****** Object:  View dbo.qryELPDetails    Script Date: 25/02/2002 16:35:38 ******/
/****** Object:  View dbo.qryELPDetails    Script Date: 01/06/2001 19:10:38 ******/
CREATE VIEW [dbo].[qryELPDetails]
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
	dbo.tblLeavePeriod.shareLeaveWithTeamCalendar, 
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
/****** Object:  View [dbo].[qryLeavePeriodDetails]    Script Date: 4/19/2018 9:29:56 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 02/05/2002 14:23:57 ******/
/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 18/04/2002 14:09:01 ******/
/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 25/02/2002 16:35:38 ******/
/****** Object:  View dbo.qryLeavePeriodDetails    Script Date: 01/06/2001 19:10:38 ******/
CREATE VIEW [dbo].[qryLeavePeriodDetails]
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
	dbo.tblLeavePeriod.shareLeaveWithTeamCalendar,
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
ALTER TABLE [dbo].[tblELPRelief]  WITH CHECK ADD  CONSTRAINT [FK_tblELPRelief_tblELPInstance] FOREIGN KEY([lngELPID])
REFERENCES [dbo].[tblELPInstance] ([lngID])
GO
ALTER TABLE [dbo].[tblELPRelief] CHECK CONSTRAINT [FK_tblELPRelief_tblELPInstance]
GO
ALTER TABLE [dbo].[tblLeavePeriod]  WITH CHECK ADD  CONSTRAINT [FK_tblLeavePeriod_tblLeaveType] FOREIGN KEY([lngLeaveTypeID])
REFERENCES [dbo].[tblLeaveType] ([lngID])
GO
ALTER TABLE [dbo].[tblLeavePeriod] CHECK CONSTRAINT [FK_tblLeavePeriod_tblLeaveType]
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[41] 4[14] 2[13] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblUser"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 136
               Right = 271
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "WorkerPrivate"
            Begin Extent = 
               Top = 6
               Left = 309
               Bottom = 326
               Right = 869
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 24
         Width = 284
         Width = 1980
         Width = 1710
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 1440
         Alias = 900
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'Admins'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'Admins'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[36] 4[7] 2[28] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1 [50] 2 [25] 3))"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1[56] 4[18] 2) )"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "tblLeavePeriod"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 114
               Right = 228
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "EE"
            Begin Extent = 
               Top = 6
               Left = 266
               Bottom = 114
               Right = 449
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "Manager"
            Begin Extent = 
               Top = 114
               Left = 38
               Bottom = 222
               Right = 251
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
      Begin ColumnWidths = 23
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 4335
         Alias = 2580
         Table = 3345
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 4365
         Or' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'qryUserApprovalsAll'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N' = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'qryUserApprovalsAll'
GO
EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'qryUserApprovalsAll'
GO
