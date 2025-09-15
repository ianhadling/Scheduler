CREATE PROC [dbo].[uspAddAdditionalTasksToSchedule]
AS

DECLARE @TasksAdded		INT 

BEGIN
BEGIN TRAN
BEGIN TRY
	
	INSERT dbo.TaskResults	(ScheduledTaskID		,TaskExpected		,Config)
	SELECT					adt.ScheduledTaskID		,adt.TaskStart		,ISNULL(adt.Config, st.Config)
	FROM dbo.AdditionalTasks		AS adt
		JOIN ScheduledTasks			AS st ON st.id = adt.ScheduledTaskID
	WHERE AddedToScheduleOn IS NULL

	SET @TasksAdded = @@ROWCOUNT

	UPDATE dbo.AdditionalTasks
		SET AddedToScheduleOn	= GETDATE()
			,ModifiedOn			= GETDATE()
			,ModifiedBy			= SUSER_NAME()
	WHERE AddedToScheduleOn IS NULL

	IF @TasksAdded  > 0
		INSERT GeneralLog (LogSource, LogStatus, LogMessage)
		VALUES ('uspAddAdditionalTasksToSchedule', 'Info', CONVERT (VARCHAR(10), @TasksAdded) + ' tasks added to live schedule')

	COMMIT
END TRY
BEGIN CATCH
	IF @@TRANCOUNT > 0
		ROLLBACK
	
	INSERT INTO ProcedureErrors (ErrorCode			,ErrorSource			,AdditionalInformation									,ErrorDescription)
		VALUES					(ERROR_NUMBER()		,ERROR_PROCEDURE()		,'Error Line: ' + CONVERT(VARCHAR(10),ERROR_LINE())		,ERROR_MESSAGE())

END CATCH


END
GO