/*
	AGREGO EL CAMPO Cuenta COMO CLAVE Y PARA ELLO PRIMERO TENGO QUE
	DECIR QUE NO PUEDE SER NULL
*/
ALTER TABLE sufractanSA.dbo.Cuenta
	ALTER COLUMN Cuenta varchar(10) NOT NULL
GO
ALTER TABLE	sufractanSA.dbo.Cuenta
	ADD PRIMARY KEY (Cuenta)
GO

/*
	ELIMINACION DE PROCEDIMIENTOS Y FUNCIONES
*/

USE [sufractanSA]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_cuenta]
GO

/*
	CREACION DE PROCEDIMIENTOS Y FUNCIONES
*/

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

	
CREATE PROCEDURE PR_modificar_cuenta
	(@cuenta varchar(10),
	@descripcion varchar(50),
	@nivel smallint,
	@empresa smallint)
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			UPDATE sufractanSA.dbo.Cuenta
			SET Descripcion = @descripcion
				, Nivel = @nivel
				, Empresa = @empresa
			WHERE Cuenta = @cuenta
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				set @mensaje_error = @mensaje_error + 'NO SE PUDO MODIFICAR LA CUENTA.'	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO
	
CREATE PROCEDURE PR_alta_cuenta
	(@cuenta varchar(10),
	@descripcion varchar(50),
	@nivel smallint,
	@empresa smallint)
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			insert into sufractanSA.dbo.Cuenta
				values (@Cuenta, @Descripcion, @Nivel, @Empresa)
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				EXEC PR_modificar_cuenta @cuenta, @descripcion, @nivel, @empresa	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO	

CREATE PROCEDURE PR_baja_cuenta
	(@cuenta varchar(10))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			DELETE FROM sufractanSA.dbo.Cuenta
			WHERE Cuenta = @cuenta
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				set @mensaje_error = @mensaje_error + 'NO SE PUDO ELIMINAR LA CUENTA.'	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO	

