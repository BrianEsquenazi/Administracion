/*
---------------------------------------------------------------
---------------------------FUNCIONES---------------------------
---------------------------------------------------------------
*/

USE [surfactanSA]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_get_fecha_ordenable]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_get_fecha_ordenable]
GO

USE [surfactanSA]
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE FUNCTION [dbo].[FN_get_fecha_ordenable](@fecha varchar(10))
RETURNS varchar(10)
AS
BEGIN
	declare @ordenFechaInt int
	declare @ordenFechaString varchar(10)
	set @ordenFechaInt = YEAR(@fecha) * 10000 + MONTH(@fecha) * 100 + DAY(@fecha)  
	set @ordenFechaString =  convert(varchar(10), @ordenFechaInt)
	RETURN @ordenFechaString
END
GO


/*
---------------------------------------------------------------
-------------------------PROCEDIMIENTOS-------------------------
---------------------------------------------------------------
*/

USE [surfactanSA]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_provincias]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_provincias]
GO

CREATE PROCEDURE [dbo].[PR_get_provincias] 
AS

		SELECT td.codigo, td.Nombre from 
			(SELECT [Provincia] as codigo
				  ,[Nombre]
				  , 0 as orden
			  FROM [surfactanSA].[dbo].[Provincia] pr
			  where pr.Nombre not in ('Exterior', 'Sin Asignar')
			union all
			SELECT [Provincia] as codigo
				  ,[Nombre]
				  , 1 as orden
			  FROM [surfactanSA].[dbo].[Provincia] pr
			  where pr.Nombre = 'Exterior'  
			) td
		order by td.orden, td.Nombre
	RETURN
GO