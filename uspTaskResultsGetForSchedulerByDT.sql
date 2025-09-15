CREATE PROC [dbo].[uspTaskResultsGetForSchedulerByDT]
	@DateFor AS DATETIME2 = NULL
AS

BEGIN
BEGIN TRY

	EXEC uspAddAdditionalTasksToSchedule

	-- DECLARE @DateFor AS DATETIME2(2) = '2025-05-01 09:00'

	SET @DateFor = ISNULL(@DateFor, GETDATE())

	SELECT TOP 100 x.TaskId
		, x.TaskResultId
		, x.TaskName
		, x.SendTaskID
		, x.SourceFolder
		, x.SourceFilename
		, x.SourceFunction
		, x.PriorityVal
		, x.Config
		, x.TaskExpected
		, rs.EmailSubject
		, rs.TemplateLocation
		, rs.EmailBody
		, rs.EmailToList
		, rs.EmailccList
		, rs.SavePath

	FROM 
		(SELECT TaskId = st.Id
				,TaskResultId = tr.ID
				, st.SendTaskID
				, st.TaskName
				, st.SourceFolder
				, st.SourceFilename
				, st.SourceFunction
				, st.PriorityVal
				, tr.TaskExpected
				, tr.Config
				, ExpectedOrder = ROW_NUMBER() OVER (PARTITION BY st.ID ORDER BY tr.TaskExpected DESC)
		FROM dbo.TaskResults					AS tr
			JOIN dbo.ScheduledTasks				AS st	ON st.ID = tr.ScheduledTaskID
			/* 18-06-2025	Ian H	Timing for the scheduler is not accurate to the second, and so can start early. Adding 1 minute tolerance should resolve this */
			/* 08-07-2025	Ian H	Also ensure that all tasks to run are and no earlier than midnight today */
		WHERE tr.TaskExpected BETWEEN CONVERT (DATE, @DateFor) AND  DATEADD(MINUTE, 1, @DateFor)
			AND tr.TaskStart IS NULL
		)	AS x
		LEFT JOIN dbo.ReportSettings		AS rs	ON rs.ScheduledTaskID = x.TaskId

	WHERE x.ExpectedOrder = 1
	ORDER BY TaskExpected, PriorityVal

END TRY
BEGIN CATCH
	
	INSERT INTO ProcedureErrors (ErrorCode			,ErrorSource			,AdditionalInformation									,ErrorDescription)
		VALUES					(ERROR_NUMBER()		,ERROR_PROCEDURE()		,'Error Line: ' + CONVERT(VARCHAR(10),ERROR_LINE())		,ERROR_MESSAGE())

END CATCH
END
