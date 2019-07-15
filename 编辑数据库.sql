USE [CASCRS_VOUCHER]
GO
alter table [dbo].[CodeContrast] add Flag Tinyint default 1
GO
UPDATE [dbo].[CodeContrast]
   SET [Flag] = 1
 WHERE Flag is null
GO


