CREATE PROCEDURE [dbo].[uspGeneralLogAdd]
	@LogSource		VARCHAR (100)
	,@LogStatus		VARCHAR(50)
	,@LogMessage	VARCHAR(500)
	,@AddnlInfo		VARCHAR(500) = NULL

AS 
BEGIN
	INSERT GeneralLog	(LogSource	,LogStatus	,LogMessage		,AddnlInfo)
	VALUES				(@LogSource	,@LogStatus	,@LogMessage	,@AddnlInfo)

END