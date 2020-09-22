CREATE DATABASE SERVICECENTRE
USE SERVICECENTRE

CREATE TABLE [dbo].[USSC] (
	[USSC_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[ServiceNo] [varchar] (20) NULL ,
	[Dated] [varchar] (20) NULL ,
	[NOC] [varchar] (70) NULL ,
	[TEL] [varchar] (50) NULL ,
	[ADDRESS] [varchar] (150) NULL ,
	[PRINTSTATUS] [varchar] (50) NULL ,
	[PRINTBILL] [varchar] (50) NULL ,
	[BILLINGPRINT] [varchar] (50) NULL ,
	[VNO] [varchar] (20) NULL ,
	[TYPEOFVEHICLE] [varchar] (50) NULL, 
	[ENTRYTIME] [varchar] (20) NULL ,
	[EXITTIME] [varchar] (20) NULL ,
	[FULLSERVICE] [varchar] (60) NULL ,
	[HALFSERVICE] [varchar] (60) NULL ,
	[UNDERSIDE] [varchar] (60) NULL ,
	[ENGINEWASHING] [varchar] (60) NULL ,
	[WATERSPRAY] [varchar] (60) NULL ,
	[GELINGWASH] [varchar] (60) NULL ,
	[AMOUNT] [varchar] (20) NULL ,
	[PAYMENTSTATUS] [varchar] (20) NULL ,
	[PAYMENTDATE] [varchar] (20) NULL ,
	[BILLNO] [varchar] (20) NULL ,
	[PAYMENTRECNO] [varchar] (10) NULL ,
	[PAYMENTRECBY] [varchar] (50) NULL ,
	[FULLSERVICEAMT] [varchar] (10) NULL ,
	[HalfServiceAmt] [varchar] (10) NULL,
	[JALFSERVICEAMT] [varchar] (10) NULL ,
	[UNDERSIDEAMT] [varchar] (10) NULL ,
	[ENGINEWASHINGAMT] [varchar] (10) NULL ,
	[WATERSPRAYAMT] [varchar] (10) NULL ,
	[GELINGWASHAMT] [varchar] (10) NULL 
) ON [PRIMARY]

CREATE TABLE [dbo].[NUMBERS] (
	[USSC_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[JobCard] [varchar] (20) NULL ,
	[BillNo] [varchar] (20) NULL ,
	[RecNo] [varchar] (70) NULL 
) ON [PRIMARY]
CREATE TABLE [dbo].[Vehicle] (
	[USSC_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[TypeofVehicle] [varchar] (20) NULL 
) ON [PRIMARY]

CREATE TABLE [dbo].[LOGIN] (
	[USSC_ID] [int] IDENTITY (1, 1) NOT NULL ,
	[USERNAME] [varchar] (20) NULL ,
	[USER] [varchar] (20) NULL ,
	[PASSWORD] [varchar] (70) NULL 
) ON [PRIMARY]


SELECT * FROM USSC

SELECT * FROM login