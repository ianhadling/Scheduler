CREATE PROCEDURE [dbo].[uspTaskResultsUpdate]
	@TaskResultID		AS INT
	,@TaskStatus		AS VARCHAR(50)
	,@Message			AS VARCHAR(500)
AS
BEGIN

SET NOCOUNT ON

	IF @TaskStatus = 'Start'
	BEGIN
		UPDATE dbo.TaskResults
			SET TaskStart = GETDATE()
		WHERE ID = @TaskResultID
	END

	ELSE IF @TaskStatus IN ('Success', 'Fail')
		UPDATE dbo.TaskResults
			SET TaskEnd = GETDATE()
				, TaskCompleted = IIF(@TaskStatus = 'Success', 1, 0)
		WHERE ID = @TaskResultID
END
