/*****************************************************************************************
Name:			uspTaskResultsAddNewTasksByDateRange
Description:	Updates table TaskResults with all expected schedule tasks runs for the given period. This is repeatable to allow new scheduled tasks to be added.
				Tasks can have single or multiple schedules, with or without recurring runs per day. This will only add future tasks (see end of proc), and so can add
				tomorrow morning's scheduled items, ignoring them for today
Request Source:	Report/Task rewrite project
Author:			Ian Hadlington
Created on:		09-Jun-2025
Called From:	SQL Server Agent on a weekly, and ad-hoc basis
Parameters:		@DateFrom			-- Given period for which the Task Results schedule is to be updated for
				@DateTo

Step by Step breakdown
======================
		Single Schedules
	1. Add all daily scheduled tasks with no recurrences, and a single schedule
	2. Add all daily scheduled tasks with recurring runs, and a single schedule
	3. Add all weekday scheduled tasks with no recurrences, and a single schedule
	4. Add all weekday scheduled tasks with recurring runs, and a single schedule
	5. Add first work day of the week scheduled tasks with no recurrences, and a single schedule
	6. Add first work day of the month scheduled tasks with no recurrences, and a single schedule
	7. Add third work day of the month scheduled tasks with no recurrences, and a single schedule
	8. Add third working Monday of the month scheduled tasks with no recurrences, and a single schedule


		Multiple Schedules

Change Control
==============

When		Who		Version		What
====		===		=======		====
18-06-2025	Ian H	v1.01		Remove the section that adds tasks from the AdditionalTasks table. This is now added to uspTaskResultsGetForSchedulerByDT
24-06-2025	Ian H	v1.02		Add config to TaskResults table from ScheduledTasks table to allow configs to change, and to be differente for adhoc (often repeated) tasks
04-08-2025	Ian H	v1.03		Added multiple schedules to 6,7,8. Also combined 6 and 7 (allowing Nth working day of month, instead of just first and third)
26-08-2025	Ian H	v1.04		Modified logic regarding working days. Can now schedule daily tasks on bank holidays or not.

*****************************************************************************************/

CREATE PROC [dbo].[uspTaskResultsAddNewTasksByDateRange]
	@DateFrom		AS DATE
	,@DateTo		AS DATE = NULL

AS

BEGIN

	SET NOCOUNT ON 
	/*
	DECLARE @DateFrom	DATE = '2025-05-01'
	DECLARE @DateTo		DATE = '2025-05-31'
	*/
BEGIN TRY
	/* Set @DateTo to one week in the future if it is NULL */
	SET @DateTo = ISNULL(@DateTo, DATEADD(DAY, 7, @DateFrom));

	SET DATEFIRST 1;

	DROP TABLE IF EXISTS #NewSchedule;

	CREATE TABLE #NewSchedule (
		ScheduledTaskID		INT
		,Config				XML
		,TaskScheduleID		INT
		,TaskExpected		DATETIME2(2)
		)

	-- Add Daily tasks			

	-- 	1. Add all daily scheduled tasks with no recurrence, and single or multiple schedules

	INSERT #NewSchedule (ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT				st.ID				,ts.ID						, CAST (wd.DateCal AS DATETIME) + CAST (ts.StartAt AS DATETIME)
	FROM dbo.ScheduledTasks							AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS JOIN 
			(SELECT c.DateCal, bh.bhDate		FROM dbo.vwCalendar			AS c															/* 26-08-2025	Ian H	v1.04 */
				LEFT JOIN dbo.VwBankHolidays	AS bh		ON c.DateCal = bh.bhDate
			WHERE c.DateCal BETWEEN @DateFrom AND @DateTo
				AND c.IsWeekend = 0
			)		wd
	WHERE ts.PeriodType = 1			-- Or could add link to TaskSchedules and used this
		AND COALESCE (ts.IsMonday, ts.IsTuesday, ts.IsWednesday, ts.IsThursday, ts.IsFriday, 0) = 0
		AND st.IsActive = 1
		AND ((ts.WorkingDay = 1 AND bhDate IS NULL)																							/* 26-08-2025	Ian H	v1.04 */
			OR ISNULL(ts.WorkingDay,0) = 0) 
		AND ts.IsRecurring = 0
	;

	-- 	2. Add all daily scheduled tasks with recurring runs, and a single schedule

	INSERT #NewSchedule ( ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT			 	st.ID				,ts.ID						, CAST (wd.DateCal AS DATETIME) + CAST (rec.TimeValue AS DATETIME)
	FROM dbo.ScheduledTasks						AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS APPLY dbo.ufnGetRecurringTimes(ts.StartAt, ts.RecurringTo, ts.RecurringInterval) AS rec
		CROSS JOIN 
			(SELECT c.DateCal, bh.bhDate		FROM dbo.vwCalendar			AS c															/* 26-08-2025	Ian H	v1.04 */						LEFT JOIN dbo.VwBankHolidays	AS bh		ON c.DateCal = bh.bhDate
			WHERE c.DateCal BETWEEN @DateFrom AND @DateTo
				AND c.IsWeekend = 0
			)		wd
	WHERE 1=1
		AND ts.IsRecurring = 1
		AND ts.PeriodType = 1
		AND COALESCE (ts.IsMonday, ts.IsTuesday, ts.IsWednesday, ts.IsThursday, ts.IsFriday, 0) = 0
		AND st.IsActive = 1
		AND ((ts.WorkingDay = 1 AND bhDate IS NULL)																							/* 26-08-2025	Ian H	v1.04 */
			OR ISNULL(ts.WorkingDay,0) = 0) 


	-- 3. Add all weekday scheduled tasks with no recurrence, and single/multiple schedule(s)

	INSERT #NewSchedule ( ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT				st.ID				,ts.ID						, CAST (wd.DateCal AS DATETIME) + CAST (ts.StartAt AS DATETIME)
	FROM dbo.ScheduledTasks			AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS JOIN 
			(SELECT DateCal, c.WeekCalday		FROM dbo.vwCalendar			AS c
			WHERE c.DateCal BETWEEN @DateFrom AND @DateTo
				AND c.IsWeekend = 0
			)		wd
	WHERE ts.IsRecurring = 0
		AND st.IsActive = 1
		AND ts.PeriodType = 2
		AND CASE DATEPART (WEEKDAY, wd.DateCal)
				WHEN 1 THEN ts.IsMonday
				WHEN 2 THEN ts.IsTuesday
				WHEN 3 THEN ts.IsWednesday
				WHEN 4 THEN ts.IsThursday
				WHEN 5 THEN ts.IsFriday
				ELSE 0
			END = 1
	;

	-- 4. Add all weekday scheduled tasks with recurring runs, and single/multiple schedule(s)

	INSERT #NewSchedule ( ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT				st.ID				,ts.ID						, CAST (wd.DateCal AS DATETIME) + CAST (rec.TimeValue AS DATETIME)
	FROM dbo.ScheduledTasks			AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS APPLY dbo.ufnGetRecurringTimes(ts.StartAt, ts.RecurringTo, ts.RecurringInterval)		AS rec
		CROSS JOIN 
			(SELECT DateCal, c.WeekCalday		FROM dbo.vwCalendar			AS c
			WHERE c.DateCal BETWEEN @DateFrom AND @DateTo
				AND c.IsWeekend = 0
			)		wd
	WHERE ts.IsRecurring = 1
		AND rec.TimeValue IS NOT NULL
		AND st.IsActive = 1
		AND ts.PeriodType = 2
		AND CASE DATEPART (WEEKDAY, wd.DateCal)
				WHEN 1 THEN ts.IsMonday
				WHEN 2 THEN ts.IsTuesday
				WHEN 3 THEN ts.IsWednesday
				WHEN 4 THEN ts.IsThursday
				WHEN 5 THEN ts.IsFriday
				ELSE 0
			END = 1
	;

	-- 	5. Add Nth work day of the week scheduled tasks with no recurrences, and a single/multiple schedule

	INSERT #NewSchedule ( ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT				st.ID				,ts.ID						, CAST (wdw.CalDate AS DATETIME) + CAST (ts.StartAt AS DATETIME)
	FROM dbo.ScheduledTasks			AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS APPLY dbo.ufn_GetWorkingDaysInWeek(@DateFrom, @DateTo)		AS wdw
	WHERE ts.IsRecurring = 0
		AND st.IsActive = 1
		AND ts.PeriodType = 2
		AND ts.NthPerPeriod = wdw.WorkDayNumber
		AND ts.WorkingDay = 1
	;

	-- 	6. Add Nth work day of the month scheduled tasks with no recurrences, and a single/multiple schedule

	INSERT #NewSchedule ( ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT				st.ID				,ts.ID						, CAST (wdm.CalDate AS DATETIME) + CAST (ts.StartAt AS DATETIME)
	FROM dbo.ScheduledTasks			AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS APPLY dbo.ufn_GetWorkingDaysInMonth(YEAR(@DateFrom))		AS wdm
	WHERE ts.IsRecurring = 0
		AND st.IsActive = 1
		AND ts.PeriodType = 3
		AND ts.NthPerPeriod = wdm.DayNumber					/*	04-08-2025	Ian H	v1.03	*/
		AND ts.WorkingDay = 1
		AND wdm.CalDate BETWEEN @DateFrom AND @DateTo
	;
/*
	-- 7. Add third work day of the month scheduled tasks with no recurrences, and a single/multiple schedule
	INSERT #NewSchedule ( ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT				st.ID				,ts.ID						, CAST (wdm.CalDate AS DATETIME) + CAST (ts.StartAt AS DATETIME)
	FROM dbo.ScheduledTasks			AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS APPLY dbo.ufn_GetWorkingDaysInMonth(YEAR(@DateFrom))		AS wdm
	WHERE ts.IsRecurring = 0
		AND st.IsActive = 1
		AND ts.PeriodType = 3
		AND ts.NthPerPeriod = 3
		AND ts.WorkingDay = 1
		AND wdm.DayNumber = 3
		AND wdm.CalDate BETWEEN @DateFrom AND @DateTo
	;
*/
	-- 8. Add First working Monday of the month scheduled tasks with no recurrences, and a single schedule

	INSERT #NewSchedule ( ScheduledTaskID	,TaskScheduleID				, TaskExpected)
	SELECT				st.ID				,ts.ID						, CAST (wdm.CalDate AS DATETIME) + CAST (ts.StartAt AS DATETIME)
	FROM dbo.ScheduledTasks			AS st
		LEFT JOIN dbo.ScheduledTasks_TaskSchedules	AS stts	ON stts.ScheduledTaskID = st.id
															AND stts.IsActive = 1
		JOIN dbo.TaskSchedules						AS ts	ON ts.id IN (stts.TaskScheduleID, st.TaskScheduleID)
		CROSS APPLY dbo.ufn_GetNthWeekdayOfMonth(YEAR(@DateFrom))		AS wdm
	WHERE ts.IsRecurring = 0
		AND st.IsActive = 1
		AND ts.PeriodType = 3
		AND ts.NthPerPeriod = 1
		AND ts.IsMonday = 1
		AND wdm.WeekCalday = 1				-- Get Monday
		AND wdm.NthOfTheMonth = 1			-- Get 1st Monday
		AND wdm.CalDate BETWEEN @DateFrom AND @DateTo
	;

	/*
	select top 100 * from vwcalendar
	*/

	UPDATE ns
		SET Config = st.Config
	FROM #NewSchedule			AS ns
		JOIN dbo.ScheduledTasks	AS st ON st.id = ns.ScheduledTaskID
	WHERE st.Config IS NOT NULL

	DECLARE @AddedTasks AS INT

	-- Update Expected tasks adding any missing tasks

	INSERT dbo.TaskResults (ScheduledTaskID		,Config		,TaskScheduleID		,TaskExpected)
	SELECT					n.ScheduledTaskID	,n.Config	,n.TaskScheduleID	,n.TaskExpected
	FROM #NewSchedule				AS n
		LEFT JOIN dbo.TaskResults	AS r	ON r.ScheduledTaskID	= n.ScheduledTaskID
										AND r.TaskExpected			= n.TaskExpected
	WHERE r.ScheduledTaskID IS NULL
		AND n.TaskExpected > GETDATE()
	ORDER BY 3,1

	SET @AddedTasks = @@ROWCOUNT

	IF @AddedTasks  > 0
		INSERT GeneralLog (LogSource, LogStatus, LogMessage)
		VALUES			('uspTaskResultsAddNewTasksByDateRange','Info',CONVERT (VARCHAR(10), @AddedTasks) + ' scheduled tasks added to live schedule')

	DROP TABLE IF EXISTS #NewSchedule;
END TRY
BEGIN CATCH
	
	INSERT INTO ProcedureErrors (ErrorCode			,ErrorSource			,AdditionalInformation									,ErrorDescription)
		VALUES					(ERROR_NUMBER()		,ERROR_PROCEDURE()		,'Error Line: ' + CONVERT(VARCHAR(10),ERROR_LINE())		,ERROR_MESSAGE())

END CATCH

END
GO
/****** Object:  StoredProcedure [dbo].[uspTaskResultsAddNewTasksByDateRange_Next]    Script Date: 15/09/2025 14:12:17 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
