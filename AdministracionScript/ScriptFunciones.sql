USE [surfactanSA]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_primera_cuenta]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_primera_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_ultima_cuenta]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_ultima_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_siguiente_cuenta]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_siguiente_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_anterior_cuenta]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_anterior_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_get_cuenta]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_get_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_get_banco]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_get_banco]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_get_tipo_cambio]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_get_tipo_cambio]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FN_get_rubro]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
DROP FUNCTION [dbo].[FN_get_rubro]
GO

/*
	CREACION DE FUNCIONES
*/

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION FN_get_cuenta ( @accion varchar(10), @cuenta varchar(10) )
RETURNS @cuenta_retorno TABLE
   (
	cuenta varchar(10)
	,descripcion varchar(10)
   )
AS
BEGIN
	declare @ultimo varchar(10) = (SELECT MAX(cu.Cuenta) FROM surfactanSA.dbo.Cuenta cu)
	declare @primero varchar(10) = (SELECT MIN(cu.Cuenta) FROM surfactanSA.dbo.Cuenta cu)
	IF(@accion = 'primero' or @accion = 'siguiente')
	BEGIN
		INSERT @cuenta_retorno
			SELECT cu_2.Cuenta, cu_2.Descripcion
			FROM surfactanSA.dbo.Cuenta cu_2
			WHERE cu_2.Cuenta = (SELECT TOP 1 cu_int.Cuenta 
									FROM surfactanSA.dbo.Cuenta cu_int 
									WHERE LTRIM(RTRIM(cu_int.Cuenta)) > CASE  
																			WHEN @accion = 'primero' or @ultimo <= @cuenta THEN ''
																			ELSE @cuenta
																		END  
									ORDER BY cu_int.Cuenta)
	END
		ELSE
	BEGIN
		INSERT @cuenta_retorno
			SELECT cu_1.Cuenta, cu_1.Descripcion
			FROM surfactanSA.dbo.Cuenta cu_1
			WHERE cu_1.Cuenta = (SELECT TOP 1 cu_int.Cuenta 
									FROM surfactanSA.dbo.Cuenta cu_int 
									WHERE LTRIM(RTRIM(cu_int.Cuenta)) < CASE  
																			WHEN @accion = 'ultimo' or @primero >= @cuenta THEN 'zzzzzzzzzz'
																			ELSE @cuenta
																		END 
									ORDER BY cu_int.Cuenta DESC) 
	END
	RETURN
END
GO


CREATE FUNCTION [dbo].[FN_get_banco] ( @accion varchar(10), @banco smallint )
RETURNS @banco_retorno TABLE
   (
	banco smallint
	,nombre varchar(50)
	,cuenta varchar(10)
   )
AS
BEGIN
	declare @ultimo varchar(10) = (SELECT MAX(ba.banco) FROM surfactanSA.dbo.Banco ba)
	declare @primero varchar(10) = (SELECT MIN(ba.banco) FROM surfactanSA.dbo.Banco ba)
	IF(@accion = 'primero' or @accion = 'siguiente')
	BEGIN
		INSERT @banco_retorno
			SELECT ba_2.Banco, ba_2.Nombre, ba_2.Cuenta
			FROM surfactanSA.dbo.Banco ba_2
			WHERE  ba_2.Banco = (SELECT TOP 1 ba_int.Banco 
									FROM surfactanSA.dbo.Banco ba_int 
									WHERE LTRIM(RTRIM(ba_int.Banco)) > CASE  
																			WHEN @accion = 'primero' or @ultimo <= @banco THEN (-1)
																			ELSE @banco
																		END  
									ORDER BY ba_int.Banco)
	END
		ELSE
	BEGIN
		INSERT @banco_retorno
			SELECT ba_2.Banco, ba_2.Nombre, ba_2.Cuenta
				FROM surfactanSA.dbo.Banco ba_2
				WHERE  ba_2.Banco = (SELECT TOP 1 ba_int.Banco 
										FROM surfactanSA.dbo.Banco ba_int 
										WHERE LTRIM(RTRIM(ba_int.Banco)) < CASE  
																				WHEN @accion = 'ultimo' or @primero >= @banco THEN 32767
																				ELSE @banco
																			END  
										ORDER BY ba_int.Banco DESC)
	END
	RETURN
END
GO


CREATE FUNCTION [dbo].[FN_get_tipo_cambio] ( @accion varchar(10), @fecha varchar(10) )
RETURNS @retorno TABLE
   (
	fecha varchar(10)
	,cambio  float
   )
AS 
BEGIN
	declare @OrdFecha varchar(10) = dbo.get_fecha_ordenable(@fecha)
	declare @ultimo varchar(10) = (SELECT MAX(ca.OrdFecha) FROM surfactanSA.dbo.CambioAdm ca)
	declare @primero varchar(10) = (SELECT MIN(ca.OrdFecha) FROM surfactanSA.dbo.CambioAdm ca)
	IF(@accion = 'primero' or @accion = 'siguiente')
	BEGIN
		INSERT @retorno
			SELECT ca_2.fecha, ca_2.cambio
			FROM surfactanSA.dbo.CambioAdm ca_2
			WHERE ca_2.OrdFecha = (SELECT TOP 1 ca_int.OrdFecha 
									FROM surfactanSA.dbo.CambioAdm ca_int 
									WHERE LTRIM(RTRIM(ca_int.OrdFecha)) > CASE  
																			WHEN @accion = 'primero' or @ultimo <= @OrdFecha THEN '0'
																			ELSE @OrdFecha
																		END  
									ORDER BY ca_int.OrdFecha)
	END
		ELSE
	BEGIN
		INSERT @retorno
			SELECT ca_2.fecha, ca_2.cambio
			FROM surfactanSA.dbo.CambioAdm ca_2
			WHERE ca_2.OrdFecha = (SELECT TOP 1 ca_int.OrdFecha 
									FROM surfactanSA.dbo.CambioAdm ca_int 
									WHERE LTRIM(RTRIM(ca_int.OrdFecha)) < CASE  
																			WHEN @accion = 'ultimo' or @primero >= @OrdFecha THEN '99999999'
																			ELSE @OrdFecha
																		END  
									ORDER BY ca_int.OrdFecha DESC)
	END
	RETURN
END
GO

CREATE FUNCTION [dbo].[FN_get_rubro] ( @accion varchar(10), @rubro int )
RETURNS @retorno TABLE
   (
	rubro varchar(10)
	,descripcion varchar(10)
   )
AS
BEGIN
	declare @ultimo int = (SELECT MAX(tp.Codigo) FROM surfactanSA.dbo.TipoProv tp)
	declare @primero int = (SELECT MIN(tp.Codigo) FROM surfactanSA.dbo.TipoProv tp)
	IF(@accion = 'primero' or @accion = 'siguiente')
	BEGIN
		INSERT @retorno
			SELECT tp2.Codigo, tp2.Descripcion
			FROM surfactanSA.dbo.TipoProv tp2
			WHERE tp2.Codigo = (SELECT TOP 1 tp_int.Codigo 
									FROM surfactanSA.dbo.TipoProv tp_int 
									WHERE tp_int.Codigo > CASE  
																WHEN @accion = 'primero' or @ultimo <= @rubro THEN 0
																ELSE @rubro
															END  
									ORDER BY tp_int.Codigo)
	END
		ELSE
	BEGIN
		INSERT @retorno
			SELECT tp2.Codigo, tp2.Descripcion
			FROM surfactanSA.dbo.TipoProv tp2
			WHERE tp2.Codigo = (SELECT TOP 1 tp_int.Codigo 
									FROM surfactanSA.dbo.TipoProv tp_int 
									WHERE tp_int.Codigo < CASE  
																WHEN @accion = 'ultimo' or @primero >= @rubro THEN 2000000000
																ELSE @rubro
															END 
									ORDER BY tp_int.Codigo DESC) 
	END
	RETURN
END

GO
