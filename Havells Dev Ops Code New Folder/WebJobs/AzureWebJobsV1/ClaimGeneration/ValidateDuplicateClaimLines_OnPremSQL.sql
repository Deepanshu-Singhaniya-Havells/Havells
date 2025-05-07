drop table claimjobsFeb2021
go
SELECT hil_jobidname into claimjobsFeb2021 
FROM 
(
SELECT hil_jobidname
FROM D365.hil_claimline cl 
WHERE CL.hil_claimperiodname IN ('202312')
group by hil_jobidname
) AS TAB

GO
SELECT DISTINCT ''''+hil_jobidname,hil_claimperiodname
FROM D365.hil_claimline cl 
WHERE CL.hil_claimperiodname NOT IN ('202312') AND 
hil_jobidname IN (SELECT hil_jobidname FROM claimjobsFeb2021)
GO
--DROP TABLE claimjobsFeb2021

--DROP TABLE CLAIMDUPLICATELINES
SELECT DISTINCT hil_claimheader,hil_claimcategory,hil_jobid 
INTO CLAIMDUPLICATELINES 
FROM 
(
SELECT CL.hil_claimheadername,CL.hil_franchiseename,CL.hil_claimamount,CL.hil_claimcategoryname,''''+ CL.hil_jobidname AS hil_jobidname ,
COUNT(*) as rowcnt,CL.hil_claimheader,CL.hil_claimcategory,CL.hil_jobid
FROM D365.hil_claimline CL
WHERE CL.hil_claimperiodname in ('202312')
GROUP BY CL.hil_claimheader,CL.hil_claimcategory,CL.hil_jobid,CL.hil_claimheadername,CL.hil_franchiseename,CL.hil_claimamount,CL.hil_claimcategoryname,''''+CL.hil_jobidname
HAVING COUNT(*) >1
)  AS TAB

GO
--SELECT * 
--FROM D365.Hil_Claimline CL
--INNER JOIN CLAIMDUPLICATELINES CDL ON CDL.hil_claimheader = CL.hil_claimheader AND  CDL.hil_claimcategory = CL.hil_claimcategory AND 
--CDL.hil_jobid = CL.hil_jobid

--SELECT * FROM CLAIMDUPLICATELINES
