
--Query to Retrieve Data from Trend Converstion of 1 min to 5 min
--E:\ActiveRestore\ActiveRestore_Data.MDF
--Do not delete  :: Created 04MAY2010

--Set the query to the correct database
--Note: When querying for DP points, use the PDI point NOT the TIC point.
use Active


--set variables
declare @startdate as datetime	
declare @enddate as datetime
declare @pointname as varchar(40)
declare @pointvalue as varchar(6)
declare @pointid as varchar(20)

--Set the start and end dates
select @startdate = '01/01/2000'
select @enddate = '01/02/2000'
--Set @pointname to '%' to get all points or 'ALARM' to get ALARMS
--Set @pointname to the actual pointname like 'TI200708' 
select @pointname = '%'
select @pointvalue = '%'
select @pointid = '%'

SELECT
    dbo.tblPointValue.CollectionDateTime,
    dbo.tblPointDefinition.PointName,
    dbo.tblPointValue.PointValue,
    dbo.tblPointDefinition.PointUnit,
    dbo.tblPointValue.PointId
FROM
    dbo.tblPointValue
    INNER JOIN dbo.tblPointDefinition ON dbo.tblPointValue.PointId = dbo.tblPointDefinition.PointId
WHERE   ((dbo.tblPointValue.CollectionDateTime >= @startdate AND dbo.tblPointValue.CollectionDateTime <= @enddate) 
and dbo.tblPointDefinition.PointName like '' + @pointname and dbo.tblPointValue.PointValue like ''+ @pointvalue) 
--and dbo.tblPointValue.PointId like '' + @pointid) --//added to get ALARMS if you dont comment this line out, query will fail.

 