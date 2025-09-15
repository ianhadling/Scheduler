CREATE FUNCTION [dbo].[ufnGetRecurringTimes] (@TimeFrom TIME(2), @TimeTo TIME(2), @Interval INT)
RETURNS @Times TABLE (TimeValue TIME(2))
AS
BEGIN
    WITH TimeIntervals AS (
        SELECT @TimeFrom AS TimeValue
        UNION ALL
        SELECT DATEADD (MINUTE, @Interval, TimeValue)
        FROM TimeIntervals
        WHERE DATEADD (MINUTE, @Interval, TimeValue) <= @TimeTo
    )
    INSERT INTO @Times (TimeValue)
    SELECT TimeValue
    FROM TimeIntervals
    OPTION (MAXRECURSION 100);

    RETURN;
END
GO