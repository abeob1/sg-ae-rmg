
/****** Object:  StoredProcedure [dbo].[AE_SP001_OrderChit]    Script Date: 02/17/2014 18:45:26 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- AE_SP001_OrderChit 0000000063

ALTER PROCEDURE [dbo].[AE_SP001_OrderChit]
	@OrderChitNo INT
AS
BEGIN
	SELECT T1.U_AE_Date AS [Date],
	T1.U_AE_Vtype AS [Vehicle Type],
	T1.U_AE_Gname AS [Guest Name],
	T1.U_AE_GHP AS [Guest HP],
	T1.U_AE_Vno AS [Vehicle No],
	T1.U_AE_Dname AS [Driver Name],
	T1.U_AE_Ploc AS [From],
	T1.U_AE_Dloc AS [To],
	T1.U_AE_Ono AS [Order No],
	T1.U_AE_Fno AS [Flight No],
	T1.U_AE_Ftime AS [Flight Time],
	T0.U_AE_Bcode AS [Card Code],
	T0.U_AE_Bname AS [Card Name],
	T0.U_AE_Event AS [Event],
	T0.U_AE_Order AS [Order by],
	T1.U_AE_Ptime AS [Pickup time],
	T1.U_AE_Dtime AS [Dropup time],
	T0.U_AE_Issue AS [Issued By],
	T0.U_AE_Sempl AS [Sales Employee],
	T1.U_AE_Rem1 AS [Remarks1],
	T1.U_AE_Rem2 AS [Remarks2]
	FROM [@AE_CDRIVER] T0
	INNER JOIN [@AE_CDRIVER_R] T1 ON T0.DocEntry=T1.DocEntry
	WHERE T1.U_AE_Ono=@OrderChitNo AND T1.U_AE_Stype IS NOT NULL
	
END