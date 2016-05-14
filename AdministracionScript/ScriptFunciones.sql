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


/*
	CREACION DE FUNCIONES
*/

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO


CREATE FUNCTION [dbo].[FN_primera_cuenta] ()
RETURNS @cuenta_retorno TABLE
   (
	cuenta varchar(10)
	,descripcion varchar(10)
   )
AS
BEGIN
		INSERT @cuenta_retorno
			SELECT cu_2.Cuenta, cu_2.Descripcion
			FROM surfactanSA.dbo.Cuenta cu_2
			WHERE cu_2.Cuenta = (SELECT TOP 1 cu_int.Cuenta 
									FROM surfactanSA.dbo.Cuenta cu_int 
									ORDER BY cu_int.Cuenta)
	RETURN
END

GO

CREATE FUNCTION [dbo].[FN_ultima_cuenta] ()
RETURNS @cuenta_retorno TABLE
   (
	cuenta varchar(10)
	,descripcion varchar(10)
   )
AS
BEGIN
		INSERT @cuenta_retorno
			SELECT cu_2.Cuenta, cu_2.Descripcion
			FROM surfactanSA.dbo.Cuenta cu_2
			WHERE cu_2.Cuenta = (SELECT max(cu_int.Cuenta) 
									FROM surfactanSA.dbo.Cuenta cu_int)

	RETURN
END

GO


CREATE FUNCTION FN_siguiente_cuenta ( @cuenta varchar(10) )
RETURNS @cuenta_retorno TABLE
   (
	cuenta varchar(10)
	,descripcion varchar(10)
   )
AS
BEGIN
	IF ( (SELECT MAX(cu.Cuenta) FROM surfactanSA.dbo.Cuenta cu) <> @cuenta)
		INSERT @cuenta_retorno
			SELECT cu_1.Cuenta, cu_1.Descripcion
			FROM surfactanSA.dbo.Cuenta cu_1
			WHERE cu_1.Cuenta = (SELECT TOP 1 cu_int.Cuenta 
									FROM surfactanSA.dbo.Cuenta cu_int 
									WHERE LTRIM(RTRIM(cu_int.Cuenta)) > @cuenta
									ORDER BY cu_int.Cuenta) 
	ELSE
		INSERT @cuenta_retorno
			SELECT ret.Cuenta, ret.Descripcion FROM surfactanSA.dbo.[FN_primera_cuenta] () ret
	RETURN
END

GO

CREATE FUNCTION FN_anterior_cuenta ( @cuenta varchar(10) )
RETURNS @cuenta_retorno TABLE
   (
	cuenta varchar(10)
	,descripcion varchar(10)
   )
AS
BEGIN
	IF ( (SELECT MIN(cu.Cuenta) FROM surfactanSA.dbo.Cuenta cu) <> @cuenta)
		INSERT @cuenta_retorno
			SELECT cu_1.Cuenta, cu_1.Descripcion
			FROM surfactanSA.dbo.Cuenta cu_1
			WHERE cu_1.Cuenta = (SELECT TOP 1 cu_int.Cuenta 
									FROM surfactanSA.dbo.Cuenta cu_int 
									WHERE LTRIM(RTRIM(cu_int.Cuenta)) < @cuenta
									ORDER BY cu_int.Cuenta DESC) 
	ELSE
		INSERT @cuenta_retorno
			SELECT ret.Cuenta, ret.Descripcion FROM surfactanSA.dbo.[FN_ultima_cuenta] () ret
	RETURN
END

GO

