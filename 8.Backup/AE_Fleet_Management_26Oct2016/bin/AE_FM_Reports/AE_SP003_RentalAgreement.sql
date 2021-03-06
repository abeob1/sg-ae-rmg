
/****** Object:  StoredProcedure [dbo].[AE_SP002_RentalAgreement]    Script Date: 02/17/2014 18:41:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- [AE_SP002_RentalAgreement] 10

CREATE PROCEDURE [dbo].[AE_SP003_RentalAgreement]
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
	   ,[U_AE_DName] as [Name of Driver_1]
	   ,[U_AE_Dadd] as [Driver Address_1]
	   ,[U_AE_Dcno] as [Driver Ph no_1]
	   ,U_AE_Occuption as [Occupation_1]
	   ,U_AE_Nation as [Nationality_1]
	   ,U_AE_DOB as [DOB_1]
	   ,U_AE_License as [License_1]
	   ,U_AE_Pissue as [PlaceofIssue_1]
	   ,U_AE_Exdate as [Expirydate_1]
	   ,U_AE_Passno as [PassportNo_1]
	   ,U_AE_Pissuepno as [PlaceofIssuePNo_1]
	   ,U_AE_Pexdate as [PassportExpDate_1]
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
	   ,[U_AE_dwm] as [DaysWeeksMonths]
	   ,U_AE_rate as [Rate]
	   ,U_AE_PAI as [PAIDWM]
	   ,U_AE_CDW as [CDWDWM]
	   ,U_AE_Dcfees as [DelCollFees]
	   ,U_AE_petrol as [Petrol]
	   ,U_AE_Ocharges as [OtherCharges]
	   ,U_AE_BGST as [Total]
	   ,U_AE_GST as [GST]
	   ,U_AE_Netc as [NetCharge]
	   ,U_AE_Pay as [FormofPayment]
	   ,U_AE_Deposit as [Deposit]
	   ,U_AE_Exliability as [ExcessLiability]
	   ,U_AE_charges
	FROM [@AE_SBOOKING] WHERE DocNum=@DocNum
	
END