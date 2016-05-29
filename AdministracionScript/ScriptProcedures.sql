/*
	AGREGO EL CAMPO Cuenta COMO CLAVE Y PARA ELLO PRIMERO TENGO QUE
	DECIR QUE NO PUEDE SER NULL
*/


IF NOT EXISTS (SELECT * 
			FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS 
			WHERE CONSTRAINT_TYPE = 'PRIMARY KEY' 
			AND TABLE_NAME = 'Cuenta' 
			AND TABLE_SCHEMA ='dbo' )
BEGIN
	ALTER TABLE surfactanSA.dbo.Cuenta
		ALTER COLUMN Cuenta varchar(10) NOT NULL
	
	ALTER TABLE	surfactanSA.dbo.Cuenta
		ADD PRIMARY KEY (Cuenta)
	
END
GO
/*
	ELIMINACION DE PROCEDIMIENTOS Y FUNCIONES
*/

USE [surfactanSA]
GO
/****** Object:  StoredProcedure [dbo].[PR_alta_banco]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_banco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_banco]
GO
/****** Object:  StoredProcedure [dbo].[PR_alta_cuenta]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_cuenta]
GO
/****** Object:  StoredProcedure [dbo].[PR_baja_banco]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_banco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_banco]
GO
/****** Object:  StoredProcedure [dbo].[PR_baja_cuenta]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_cuenta]
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_banco_por_codigo]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_banco_por_codigo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_banco_por_codigo]
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_banco_por_nombre]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_banco_por_nombre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_banco_por_nombre]
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_cuenta_por_codigo]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_por_codigo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cuenta_por_codigo]
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_cuenta_por_descripcion]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_por_descripcion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cuenta_por_descripcion]
GO
/****** Object:  StoredProcedure [dbo].[PR_modificar_banco]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_banco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_banco]
GO
/****** Object:  StoredProcedure [dbo].[PR_modificar_cuenta]    Script Date: 05/29/2016 16:07:53 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_cuenta]
GO
/****** Object:  StoredProcedure [dbo].[PR_modificar_cuenta]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_cuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
	
CREATE PROCEDURE [dbo].[PR_modificar_cuenta]
	(@cuenta varchar(10),
	@descripcion varchar(50),
	@nivel smallint,
	@empresa smallint)
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''''
	BEGIN TRANSACTION
		BEGIN TRY
			UPDATE surfactanSA.dbo.Cuenta
			SET Descripcion = @descripcion
				, Nivel = @nivel
				, Empresa = @empresa
			WHERE Cuenta = @cuenta
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '''') 
				set @mensaje_error = @mensaje_error + ''NO SE PUDO MODIFICAR LA CUENTA.''	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_modificar_banco]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_banco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PR_modificar_banco]
	(@banco smallint,
	@nombre varchar(50),
	@cuenta varchar(10))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''''
	BEGIN TRANSACTION
		BEGIN TRY
			UPDATE surfactanSA.dbo.Banco
			SET Nombre = @nombre
				, Cuenta = @cuenta
				, Empresa = 1
			WHERE Banco = @banco
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '''') 
				set @mensaje_error = @mensaje_error + ''NO SE PUDO MODIFICAR EL BANCO.''	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_cuenta_por_descripcion]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_por_descripcion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PR_buscar_cuenta_por_descripcion]
	(@descripcion varchar(50))
AS
	select LTRIM(RTRIM(cu.Cuenta)) as Cuenta, LTRIM(RTRIM(cu.Descripcion)) as Descripcion
	from surfactanSA.dbo.Cuenta cu
	where cu.Descripcion like ''%'' + @descripcion + ''%''
	order by cu.Descripcion
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_cuenta_por_codigo]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_por_codigo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[PR_buscar_cuenta_por_codigo]
	(@cuenta varchar(10))
AS
	SELECT LTRIM(RTRIM(cu.Cuenta)) as Cuenta
		, LTRIM(RTRIM(cu.Descripcion)) as Descripcion 
	FROM Cuenta cu 
	WHERE Cuenta = @cuenta
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_banco_por_nombre]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_banco_por_nombre]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PR_buscar_banco_por_nombre]
	(@nombre varchar(50))
AS
	select LTRIM(RTRIM(ban.Banco)) as Banco, LTRIM(RTRIM(ban.Nombre)) as Nombre, LTRIM(RTRIM(ban.Cuenta)) as Cuenta
	from surfactanSA.dbo.Banco ban
	where ban.Nombre like ''%'' + @nombre + ''%''
	order by ban.Nombre
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_buscar_banco_por_codigo]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_banco_por_codigo]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[PR_buscar_banco_por_codigo]
	(@banco smallint)
AS
	SELECT ba.Banco
		, LTRIM(RTRIM(ba.Nombre)) as Nombre
		, LTRIM(RTRIM(ba.Cuenta)) as Cuenta
	FROM Banco ba 
	WHERE Banco = @banco
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_baja_cuenta]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_cuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[PR_baja_cuenta]
	(@cuenta varchar(10))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''''
	BEGIN TRANSACTION
		BEGIN TRY
			DELETE FROM surfactanSA.dbo.Cuenta
			WHERE Cuenta = @cuenta
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '''') 
				set @mensaje_error = @mensaje_error + ''NO SE PUDO ELIMINAR LA CUENTA.''	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_baja_banco]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_banco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[PR_baja_banco]
	(@banco smallint)
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''''
	BEGIN TRANSACTION
		BEGIN TRY
			DELETE FROM surfactanSA.dbo.Banco
			WHERE Banco = @banco
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '''') 
				set @mensaje_error = @mensaje_error + ''NO SE PUDO ELIMINAR EL BANCO.''	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_alta_cuenta]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_cuenta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'	
CREATE PROCEDURE [dbo].[PR_alta_cuenta]
	(@cuenta varchar(10),
	@descripcion varchar(50),
	@nivel smallint,
	@empresa smallint)
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''''
	BEGIN TRANSACTION
		BEGIN TRY
			insert into surfactanSA.dbo.Cuenta
				values (@Cuenta, @Descripcion, @Nivel, @Empresa)
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
			EXEC PR_modificar_cuenta @cuenta, @descripcion, @nivel, @empresa	
	END CATCH
' 
END
GO
/****** Object:  StoredProcedure [dbo].[PR_alta_banco]    Script Date: 05/29/2016 16:07:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_banco]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[PR_alta_banco]
	(@banco smallint,
	@nombre varchar(50),
	@cuenta varchar(10))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''''
	BEGIN TRANSACTION
		BEGIN TRY
			insert into surfactanSA.dbo.Banco
				values (@banco, @nombre, @cuenta, 1)
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
			EXEC PR_modificar_cuenta @banco, @nombre, @cuenta
	END CATCH
' 
END
GO
