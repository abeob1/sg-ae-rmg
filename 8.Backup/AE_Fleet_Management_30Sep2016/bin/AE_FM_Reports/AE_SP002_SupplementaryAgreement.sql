
/****** Object:  StoredProcedure [dbo].[AE_SP002_SupplementaryAgreement]    Script Date: 02/17/2014 18:46:14 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- AE_SP001_OrderChit 2

ALTER PROCEDURE [dbo].[AE_SP002_SupplementaryAgreement]
	@DocNum INT
AS

BEGIN

	SELECT  [U_AE_Bname] AS [TO BE PAID BY]
	   ,[U_AE_Address] AS [Address]
	   ,U_AE_Cno as [TEL]
	   ,[U_AE_DName] as [Name of Driver]
	   ,[U_AE_Dadd] as [Driver Address]
	   ,[U_AE_Dcno] as [Driver Ph no]
	   ,U_AE_Occuption as [Occupation]
	   ,U_AE_Nation as [Nationality]
	   ,U_AE_DOB as [DOB]
	   ,U_AE_License as [License]
	   ,U_AE_Pissue as [PlaceofIssue]
	   ,U_AE_Exdate as [Expirydate]
	   ,U_AE_Passno as [PassportNo]
	   ,U_AE_Pissuepno as [PlaceofIssuePNo]
	   ,U_AE_Pexdate as [PassportExpDate]
	   ,U_AE_Vmodel AS [VehicleModel]
	   ,U_AE_Vregno as [VehRegNo]
	   ,U_AE_expecD as [DateTimeExptoReturn]
	   ,U_AE_Vexten as [Extension]
	   ,U_AE_Vout as [VehicleOut]
	   ,U_AE_Vin as [VehicleIn]
	   ,U_AE_Vkmin as [KM IN]
	   ,U_AE_Vkmout AS [KM OUT]
	   ,U_AE_Vdatein as [DateTimeIn]
	   ,U_AE_Vdatetout as [DateTimeOut]
	   ,U_AE_SPD as [SurchargePerDay]
	   ,U_AE_SPT as [Total]
	   ,U_AE_SPGST as [GST]
	   ,U_AE_SPNET as [SPNET]
	   ,U_AE_SPLIB as [ExcessLiabilityMY]
	   ,U_AE_SPRemarks Remarks
	   ,U_AE_Percode as [SA PreparedBy]
	   ,U_AE_Invcode as [SA InvoicedBy]
FROM [@AE_SBOOKING] WHERE DocNum=@DocNum
	
	
END