
create proc [dbo].[VehicleTrackingLiveReport_Web]  
@Par1 as varchar(100),
 @Par2 as varchar(100),
 @Par3 as varchar(100),
 @Par4 as varchar(100),
 @Par5 as varchar(100),
 @Par6 as varchar(100),
 @Par7 as varchar(100),
 @Par8 as varchar(100)
as  
begin  
DECLARE @SQL varchar(8000)
  
CREATE TABLE #VTable ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc1 nvarchar(250),U_AE_Loc2 nvarchar(250),U_AE_Loc3 nvarchar(250),  
   U_AE_Loc4 nvarchar(250),U_AE_Loc5 nvarchar(250),U_AE_Loc6 nvarchar(250),U_AE_Loc7 nvarchar(250),U_AE_Loc8 nvarchar(250))  
DECLARE @Loc1 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
DECLARE @Loc2 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
DECLARE @Loc3 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
DECLARE @Loc4 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
DECLARE @Loc5 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
--DECLARE @Loc6 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
DECLARE @Loc7 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
DECLARE @Loc8 TABLE ([ID] [int] IDENTITY(1,1) NOT NULL,U_AE_Loc NVARCHAR(250))  
  
INSERT INTO #VTable SELECT '','','','','','','','' FROM [dbo].[@AE_VTRACK] where u_ae_stat='O'  
  
INSERT INTO @Loc1 SELECT T0.[U_AE_Vno]     
      FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc 1' order by T0.[U_AE_Date],T0.[U_AE_Time]  
   
INSERT INTO @Loc2 SELECT T0.[U_AE_Vno]     
  FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc 2' order by T0.[U_AE_Date],T0.[U_AE_Time]  
  
INSERT INTO @Loc3 SELECT T0.[U_AE_Vno]    
  FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc 3' order by T0.[U_AE_Date],T0.[U_AE_Time]  
  
INSERT INTO @Loc4 SELECT T0.[U_AE_Vno]     
  FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc 4' order by T0.[U_AE_Date],T0.[U_AE_Time]  
    
INSERT INTO @Loc5 SELECT T0.[U_AE_Vno]    
  FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc 5' order by T0.[U_AE_Date],T0.[U_AE_Time]  
    
--INSERT INTO @Loc6 SELECT T0.[U_AE_Vno]   
--  FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc R' order by T0.[U_AE_Date],T0.[U_AE_Time]  
  
INSERT INTO @Loc7 SELECT T0.[U_AE_Vno]    
  FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc C' order by T0.[U_AE_Date],T0.[U_AE_Time]  
  
INSERT INTO @Loc8 SELECT T0.[U_AE_Vno]   
  FROM [dbo].[@AE_VTRACK]  T0 where T0.[U_AE_stat] = 'O' AND T0.[U_AE_Loc1]='Loc D'  order by T0.[U_AE_Date],T0.[U_AE_Time]   
  
UPDATE T0   
 SET T0.U_AE_Loc1 = T1.U_AE_Loc,  
  T0.U_AE_Loc2 = T2.U_AE_Loc,  
  T0.U_AE_Loc3 = T3.U_AE_Loc,  
  T0.U_AE_Loc4 = T4.U_AE_Loc,  
  T0.U_AE_Loc5 = T5.U_AE_Loc,  
  --T0.U_AE_Loc6 = T6.U_AE_Loc,  
  T0.U_AE_Loc7 = T7.U_AE_Loc,  
  T0.U_AE_Loc8 = T8.U_AE_Loc  
FROM  
    #VTable T0  
LEFT JOIN @Loc1 T1 ON T0.ID = T1.ID  
LEFT JOIN @Loc2 T2 ON T0.ID = T2.ID  
LEFT JOIN @Loc3 T3 ON T0.ID = T3.ID  
LEFT JOIN @Loc4 T4 ON T0.ID = T4.ID  
LEFT JOIN @Loc5 T5 ON T0.ID = T5.ID  
--LEFT JOIN @Loc6 T6 ON T0.ID = T6.ID  
LEFT JOIN @Loc7 T7 ON T0.ID = T7.ID  
LEFT JOIN @Loc8 T8 ON T0.ID = T8.ID  
  
   
 set @SQL = 'SELECT U_AE_Loc1 AS [' + @par1 +  
   '],U_AE_Loc2 AS [' + @Par2 +  
   '],U_AE_Loc3 AS [' + @Par3 +  
   '],U_AE_Loc4 AS [' + @Par4 +  
   '],U_AE_Loc5 AS [' + @Par5 +  
   '],U_AE_Loc7 AS [' + @Par7 +  
   '],U_AE_Loc8 AS [' + @Par8 +
  '] FROM #VTable'
    
 execute(@SQL)
 
 DELETE FROM #VTable Where U_AE_Loc1 IS NULL AND U_AE_Loc2 IS NULL AND U_AE_Loc3 IS NULL AND U_AE_Loc4 IS NULL  
  AND U_AE_Loc5 IS NULL AND U_AE_Loc6 IS NULL AND U_AE_Loc7 IS NULL AND U_AE_Loc8 IS NULL 
    
end

