'This query get the latest (greather ASCII number) transmittal for each document number'
'is useful when you have a list that contains all the issues for a set of documents and'
'you want only the latest one for each.'
SELECT T_TRACK.[H doc no], MAX(T_TRACK.[Transmittal no]) AS [Transmittal no]
FROM T_TRACK
GROUP BY [H doc no];
