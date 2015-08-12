USE [FleetManagement]
GO
/****** Object:  StoredProcedure [dbo].[Commision_ReportLess60]    Script Date: 26/5/2014 8:26:59 AM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[Commision_ReportLess60] --Commision_ReportLess60 '20140501','20140528' --
@date1 as datetime,
@date2 as datetime

as

create table #Tmp (RAno varchar(20), Company Varchar(100), RentalAmount decimal(9,2), Type_ varchar(4), PAIAmount decimal (9,2),
RefNo varchar (20), InDate date , PAYAmount decimal (9,2), INVNo varchar (20), Diff decimal(9,2), Commission decimal (9,2),
CommisionCost decimal (9,2), CreditNumber varchar(20), CreditNoteAmt decimal(9,2), CreditPerAmount decimal(9,2), ComissionAmt decimal(9,2) , PAI decimal(9,2), dayDiff numeric (6))

Declare @Docnum as varchar(20)
Declare @DocDate as date
Declare @SumApplied as decimal(9,2)
Declare @InvDocnum as varchar(20)
Declare @NunAtCard as Varchar(20)
declare @DocTotal as decimal(9,2)
Declare @InvDate as date

Declare CommisionReport cursor for
SELECT T0.[DocNum], T0.[DocDate],  T1.[SumApplied], T2.[DocNum], T2.[NumAtCard],  (T2.DocTotal - T2.Vatsum) as 'DocTotal' , 
T2.DocDate FROM ORCT T0  
INNER JOIN RCT2 T1 ON T0.DocEntry = T1.DocNum inner join  OINV T2 on T2.Docentry = T1.DocEntry 
WHERE T0.[DocDate]  >= @date1  and  T0.[DocDate]  <= @date2  
GROUP BY T0.[DocNum], T0.[DocDate], T1.[SumApplied], T2.[DocNum], T2.[NumAtCard], T2.DocTotal , T2.Vatsum, T2.DocDate
open CommisionReport

fetch next from CommisionReport into @Docnum, @DocDate, @SumApplied, @InvDocnum, @NunAtCard , @DocTotal, @InvDate


while (@@FETCH_STATUS = 0)

  begin
  
    Declare @BookingDocNum varchar(20)
    Declare @BookingType varchar(20)
    Declare @Company varchar(200)
    Declare @RentalAmount decimal(9,2)
    Declare @PAI decimal(9,2)
    Declare @Type varchar(5)
    Declare @Percentage decimal(9,2)
    Declare @PerAmount decimal(9,2)
    declare @CreditNo varchar(20)
    Declare @CreditAmount decimal(9,2)
  
    set @BookingType = @NunAtCard 
    if left(@BookingType,2) = 'SD'
     begin
      SELECT @BookingDocNum = [DocNum], @Company = [U_AE_Bname],  @PAI = [U_AE_PAI], @Type = 'SD' , @RentalAmount = case when [U_AE_Term] = 1 then  U_AE_Rate else U_AE_BGST end , @Percentage = 1.75, @PerAmount = case when [U_AE_Term] = 1 then  U_AE_Rate * 0.0175 else U_AE_BGST * 0.0175 end  FROM [dbo].[@AE_SBOOKING]  WHERE [DocNum] = SUBSTRING(@BookingType,3,LEN(@BookingType)-2)
      SELECT @CreditNo = T1.[DocNum], @CreditAmount = T0.[LineTotal] FROM RIN1 T0  INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry WHERE T0.[BaseRef] = @InvDocnum
      insert into #tmp values(@BookingDocNum,@Company, @RentalAmount, @Type, @PAI , @Docnum, @DocDate, @SumApplied, @InvDocnum,
          isnull(@RentalAmount,0) - isnull(@SumApplied,0)  , @Percentage , @PerAmount , isnull(@CreditNo,''), isnull(@CreditAmount,0.0), isnull(@CreditAmount * (@Percentage/100),0.0), isnull(@PerAmount,0.0) - isnull((@CreditAmount * (@Percentage/100)),0.0), @PAI, DATEDIFF(DAY ,@InvDate, @DocDate) )
    end
   
    else
    
     begin
	
      Declare CD_Commision cursor for
      SELECT T1.[DocNum], T1.[U_AE_Bname], 'CD' , '0.0' , 
      (select case when TX.U_AE_employee is null then 1.75 else 8.0 end from [@AE_DRIVERM] TX where TX.U_AE_Dcode = T0.[U_AE_Dcode] ) ,
      (select case when TX.U_AE_employee is null then T2.[LineTotal] * 0.0175 else (T2.[LineTotal] - isnull(T0.[U_AE_tcost],0) )  * 0.08 end from [@AE_DRIVERM] TX where TX.U_AE_Dcode = T0.[U_AE_Dcode] ) ,
      (select case when TX.U_AE_employee is null then T2.[LineTotal] else (T2.[LineTotal] - isnull(T0.[U_AE_tcost],0) )  end from [@AE_DRIVERM] TX where TX.U_AE_Dcode = T0.[U_AE_Dcode] ) FROM [dbo].[@AE_CDRIVER_R]  T0  inner join [dbo].[@AE_CDRIVER]  T1 on T0.Docentry = T1.Docentry 
       inner join  INV1 T2 on T0.U_AE_Invno = T2.DocEntry and T2.U_AE_OCNO = T0.U_AE_Ono INNER JOIN OINV T3 ON T2.DocEntry = T3.DocEntry WHERE 
       T1.[DocNum] =  SUBSTRING(@BookingType,3,LEN(@BookingType)-2)
      open CD_Commision
    
      fetch next from CD_Commision into @BookingDocNum, @Company, @Type, @PAI, @Percentage, @PerAmount, @RentalAmount
    
      while (@@FETCH_STATUS = 0)

       begin 
        SELECT @CreditNo = T1.[DocNum], @CreditAmount = T0.[LineTotal] FROM RIN1 T0  INNER JOIN ORIN T1 ON T0.DocEntry = T1.DocEntry WHERE T0.[BaseRef] = @InvDocnum
        insert into #tmp values(@BookingDocNum,@Company, @RentalAmount, @Type, @PAI , @Docnum, @DocDate, @SumApplied, @InvDocnum,
          isnull(@RentalAmount,0) - isnull(@SumApplied,0)  , @Percentage , @PerAmount , isnull(@CreditNo,''), isnull(@CreditAmount,0.0), isnull(@CreditAmount * (@Percentage/100),0.0), isnull(@PerAmount,0.0) - isnull((@CreditAmount * (@Percentage/100)),0.0), @PAI, DATEDIFF(DAY ,@InvDate, @DocDate) )

        fetch next from CD_Commision into @BookingDocNum, @Company, @Type, @PAI, @Percentage, @PerAmount, @RentalAmount
      end
      close CD_Commision
      deallocate CD_Commision
        --SELECT @BookingDocNum = T0.[DocNum], @Company = T0.[U_AE_Bname], @Type = 'CD'  FROM [dbo].[@AE_CDRIVER]  T0 WHERE T0.[DocNum] = SUBSTRING(@BookingType,3,LEN(@BookingType)-2)
    end
     fetch next from CommisionReport into @Docnum, @DocDate, @SumApplied, @InvDocnum, @NunAtCard , @DocTotal , @InvDate

  end    
close CommisionReport
deallocate CommisionReport
select * from #Tmp where dayDiff < 60
drop table #Tmp 
