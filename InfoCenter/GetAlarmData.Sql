--Query to search for alarms.

--To do: Add the ability to search by point name. TBD

use Active


--set variables
declare @startdate as datetime	
declare @enddate as datetime
declare @pointname as varchar(20)
declare @pointvalue as varchar(6)
declare @priority as varchar(6)
declare @counter as varchar(6)

--Set the start and end dates

select @startdate = '01/01/2000'
select @enddate = '01/02/2000'
--Set @pointname to '%' to get all points or 'ALARM' to get ALARMS
--Set @pointname to the actual pointname like 'TI200708' 
select @pointname = 'AIC001003'
select @priority = '%'
--select @pointvalue = '%'


SELECT
      dbo.tblPointEvent.DTTimeStamp,
      dbo.tblPointEvent.ObjectName,
      dbo.tblPointDefinition.PointDescriptor,
      dbo.tblPointEvent.ActionDescription,
      dbo.tblPointEvent.Operator,
      dbo.tblPointEvent.PointValue,
      dbo.tblPointEvent.Status 
      --dbo.tblPointEvent.EventCategoryId
      --dbo.tblPointEventComment.Comment
      -- dbo.tblPointEventComment.Comment
                                     
FROM dbo.tblPointEvent 

INNER JOIN dbo.tblPointDefinition ON dbo.tblPointEvent.PointId = dbo.tblPointDefinition.PointId

WHERE ((dbo.tblPointEvent.DTTimeStamp >= @startdate and  dbo.tblPointEvent.DTTimeStamp <= @enddate) and dbo.tblPointEvent.Status like @priority and dbo.tblPointEvent.ObjectName like @pointname)
      and (dbo.tblPointEvent.ActionDescription like 'Alarm%')
--    and dbo.tblPointEvent.ObjectName like @pointname 
--    and dbo.tblPointEvent.Status like @priority)

--select COUNT(DISTINCT dbo.tblPointEvent.ObjectName)
ORDER BY DTTimeStamp