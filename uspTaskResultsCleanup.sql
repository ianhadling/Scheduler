CREATE PROCEDURE [dbo].[uspTaskResultsCleanup] 
	@ScheduleRunTime	AS DATETIME2(2)
AS
BEGIN
	/*	Set any incomplete tasks to ended and not successfully completed, with appropriate description */
	UPDATE dbo.TaskResults
		SET TaskEnd = GETDATE()
			,TaskCompleted = 0
			,CompletionMessage	= 'Cleanup: Task started and failed to complete.'
	WHERE TaskEnd IS NULL
		AND TaskStart IS NOT NULL

	/*	Close of any repeated tasks that have missed their timeslots generally due to overrun */
	UPDATE tr
		SET TaskEnd = GETDATE()
			,TaskStart = GETDATE()
			,TaskCompleted = 0
			,CompletionMessage	= 'Cleanup: Repeated task cancelled due to subsequent run of the same task.'
	FROM dbo.TaskResults			AS tr
	WHERE tr.TaskExpected < GETDATE()
		AND tr.TaskStart IS NULL
		AND tr.ScheduledTaskID					/* Does this task have subsequent completions? */
			IN (SELECT ScheduledTaskID FROM dbo.TaskResults RunTR WHERE RunTR.TaskEnd > tr.TaskExpected)
		AND tr.TaskScheduleID					/* Is this a recurring task? */
			IN (SELECT  ID FROM dbo.TaskSchedules WHERE RecurringInterval > 0)

END
GO
