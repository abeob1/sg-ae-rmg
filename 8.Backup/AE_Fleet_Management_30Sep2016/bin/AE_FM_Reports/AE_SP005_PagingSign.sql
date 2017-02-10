-- AE_SP005_PagingSign 0000000063

ALTER PROCEDURE AE_SP005_PagingSign
	@OrderChitNo INT
AS
BEGIN
	SELECT 
	T1.U_AE_Gname AS [Guest Name]
	FROM [@AE_CDRIVER] T0
	INNER JOIN [@AE_CDRIVER_R] T1 ON T0.DocEntry=T1.DocEntry
	WHERE T1.U_AE_Ono=@OrderChitNo AND T1.U_AE_Stype IS NOT NULL
	
END