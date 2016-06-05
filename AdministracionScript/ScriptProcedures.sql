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

IF EXISTS (SELECT name FROM sysindexes WHERE name = 'IX_Banco')
	DROP INDEX IX_Banco ON surfactanSA.dbo.Banco
IF EXISTS (SELECT name FROM sysindexes WHERE name = 'IX_Banco_1')
	DROP INDEX IX_Banco_1 ON surfactanSA.dbo.Banco

IF NOT EXISTS (SELECT * 
			FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS 
			WHERE CONSTRAINT_TYPE = 'PRIMARY KEY' 
			AND TABLE_NAME = 'Banco' 
			AND TABLE_SCHEMA ='dbo' )
BEGIN
	ALTER TABLE surfactanSA.dbo.Banco
		ALTER COLUMN Banco smallint NOT NULL
	
	ALTER TABLE surfactanSA.dbo.Banco
		ADD PRIMARY KEY (Banco)
	
END
GO

IF NOT EXISTS (SELECT * 
			FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS 
			WHERE CONSTRAINT_TYPE = 'PRIMARY KEY' 
			AND TABLE_NAME = 'TipoProv' 
			AND TABLE_SCHEMA ='dbo' )
BEGIN
	ALTER TABLE surfactanSA.dbo.TipoProv
		ALTER COLUMN Codigo int NOT NULL
	
	ALTER TABLE	surfactanSA.dbo.TipoProv
		ADD PRIMARY KEY (Codigo)
	
END
GO
/*
	ELIMINACION DE PROCEDIMIENTOS 
*/

USE [surfactanSA]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_banco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_banco]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_banco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_banco]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_banco_por_codigo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_banco_por_codigo]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_banco_por_nombre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_banco_por_nombre]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_por_codigo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cuenta_por_codigo]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_por_descripcion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cuenta_por_descripcion]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_banco]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_banco]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_cuenta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_cuenta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_tipo_cambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_tipo_cambio]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_tipo_cambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_tipo_cambio]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_tipo_cambio]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_tipo_cambio]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_tipo_cambio_por_fecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_tipo_cambio_por_fecha]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_rubro_proveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_rubro_proveedor]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_rubro_proveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_rubro_proveedor]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_rubro_proveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_rubro_proveedor]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_rubro_proveedor_por_codigo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_rubro_proveedor_por_codigo]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_rubro_proveedor_por_descripcion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_rubro_proveedor_por_descripcion]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_proveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_proveedor]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_baja_proveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_baja_proveedor]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_proveedor_por_codigo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_proveedor_por_codigo]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_proveedor_por_nombre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_proveedor_por_nombre]
GO

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
			EXEC PR_modificar_banco @banco, @nombre, @cuenta
	END CATCH
' 
END
GO


CREATE PROCEDURE [dbo].[PR_modificar_tipo_cambio]
	(@fecha varchar(10),
	@paridad float,
	@fecha_ord varchar(10))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			UPDATE surfactanSA.dbo.CambioAdm
			SET Fecha = @fecha
				, Cambio = @paridad 
			WHERE OrdFecha = @fecha_ord
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				set @mensaje_error = @mensaje_error + 'NO SE PUDO MODIFICAR EL CAMBIO.'	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO	


CREATE PROCEDURE [dbo].[PR_alta_tipo_cambio]
	(@fecha varchar(10),
	@paridad float)
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION

			declare @ordenFechaString varchar(10) = dbo.FN_get_fecha_ordenable(@fecha)
		IF(NOT EXISTS (select 1 from CambioAdm cam where cam.OrdFecha = @ordenFechaString))
		BEGIN			
			insert into surfactanSA.dbo.CambioAdm
				values (@fecha, @paridad, @ordenFechaString)
		END
			ELSE
		BEGIN
			EXEC PR_modificar_tipo_cambio @fecha, @paridad, @ordenFechaString
		END
	IF @@ERROR = 0 COMMIT TRANSACTION
GO

CREATE PROCEDURE [dbo].[PR_baja_tipo_cambio]
	(@fecha varchar(10))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			DELETE FROM surfactanSA.dbo.CambioAdm
			WHERE Fecha = @fecha
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				set @mensaje_error = @mensaje_error + 'NO SE PUDO ELIMINAR EL CAMBIO.'	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO

CREATE PROCEDURE [dbo].[PR_buscar_tipo_cambio_por_fecha]
	(@fecha varchar(10))
AS
	SELECT ca.Fecha
		, ca.Cambio
	FROM surfactanSA.dbo.CambioAdm ca
	WHERE ca.Fecha = @fecha

GO

CREATE PROCEDURE [dbo].[PR_modificar_rubro_proveedor]
	(@codigo int,
	@descripcion char(50))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			UPDATE surfactanSA.dbo.TipoProv
			SET Descripcion = @descripcion
			WHERE Codigo = @codigo
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				set @mensaje_error = @mensaje_error + 'NO SE PUDO MODIFICAR EL RUBRO.'	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO

CREATE PROCEDURE [dbo].[PR_alta_rubro_proveedor]
	(@codigo int,
	@descripcion char(50))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			insert into surfactanSA.dbo.TipoProv
				values (@codigo, @descripcion)
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
			EXEC PR_modificar_rubro_proveedor @codigo, @descripcion
	END CATCH
GO

CREATE PROCEDURE [dbo].[PR_baja_rubro_proveedor]
	(@codigo int)
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			DELETE FROM surfactanSA.dbo.TipoProv
			WHERE Codigo = @codigo
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				set @mensaje_error = @mensaje_error + 'NO SE PUDO ELIMINAR EL RUBRO.'	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO

CREATE PROCEDURE [dbo].[PR_buscar_rubro_proveedor_por_codigo]
	(@codigo int)
AS
	SELECT tp.Codigo
		, LTRIM(RTRIM(tp.Descripcion)) as Descripcion
	FROM TipoProv tp 
	WHERE tp.Codigo = @codigo
GO

CREATE PROCEDURE [dbo].[PR_buscar_rubro_proveedor_por_descripcion]
	(@descripcion varchar(50))
AS
	select tp.Codigo as Codigo, LTRIM(RTRIM(tp.Descripcion)) as Descripcion
	from surfactanSA.dbo.TipoProv tp
	where tp.Descripcion like '%' + @descripcion + '%'
	order by tp.Descripcion
GO

CREATE PROCEDURE [dbo].[PR_alta_proveedor]
	(@Proveedor varchar(11) ,
	@Nombre varchar(50) ,
	@Direccion varchar(50) ,
	@Localidad varchar(50) ,
	@Provincia varchar(2) ,
	@Postal varchar(4) ,
	@Region int ,
	@Telefono varchar(30) ,
	@Dias varchar(20) ,
	@Email varchar(400) ,
	@Observaciones varchar(50) ,
	@Cuit varchar(15) ,
	@Tipo varchar(1) ,
	@Iva varchar(1) ,
	@Cuenta varchar(10) ,
	@NombreCheque varchar(50) ,
	@CodIb int ,
	@CodIbCaba int ,
	@NroIb char(20) ,
	@PorceIb float ,
	@PorceIbCaba float ,
	@Rubro int ,
	@NroInsc char(15) ,
	@FechaNroInsc char(10) ,
	@CategoriaI int ,
	@CategoriaII int, 
	@FechaCategoria char(10) ,
	@IbCiudadII int,
	@Cai char(14) ,
	@VtoCai char(10) ,
	@Iso int ,
	@VtoIso char(10) ,
	@Estado int ,
	@Califica float ,
	@FechaCalifica char(10) ,
	@ObservacionesII text ,
	@Cufe char(20) ,
	@CufeII char(20) ,
	@CufeIII char(20) ,
	@DirCufe char(50) ,
	@DirCufeII char(50) ,
	@DirCufeIII char(50))
AS



	DECLARE @OrdFechaCalifica varchar(8) = (SELECT dbo.FN_get_fecha_ordenable (@FechaCalifica))
	DECLARE @OrdFechaCategoria varchar(8) = (SELECT dbo.FN_get_fecha_ordenable (@FechaCategoria))
	DECLARE @OrdFechaNroInsc varchar(8) = (SELECT dbo.FN_get_fecha_ordenable (@FechaNroInsc))
	DECLARE @Wdate char(10) = cast(getdate() AS DATE)
	--@Embargo char(1) , //ACA NUNCA IMPORTA --> S = rojo
	BEGIN TRANSACTION
		IF (NOT EXISTS (SELECT 1 FROM surfactanSA.dbo.Proveedor p WHERE p.Proveedor = @Proveedor))
					INSERT INTO surfactanSA.dbo.Proveedor (Proveedor,	Nombre, Direccion, Localidad,Provincia, Postal,Region,Telefono,Dias,
					Email,Observaciones,Cuit,Tipo,Iva,Cuenta,NombreCheque,CodIb,CodIbCaba,NroIb,
					PorceIb,PorceIbCaba,TipoProv,NroInsc,FechaNroInsc,CategoriaI,CategoriaII, FechaCategoria,
					IbCiudadII,Cai,VtoCai,Iso,VtoIso,Estado,Califica,FechaCalifica,ObservacionesII,
					Cufe,CufeII,CufeIII,DirCufe,DirCufeII,DirCufeIII, OrdFechaCalifica, OrdFechaCategoria,
					OrdFechaNroInsc, Wdate)
				VALUES (@Proveedor,	@Nombre, @Direccion, @Localidad,@Provincia, @Postal,@Region,@Telefono,@Dias,
					@Email,@Observaciones,@Cuit,@Tipo,@Iva,@Cuenta,@NombreCheque,@CodIb,@CodIbCaba,@NroIb,
					@PorceIb,@PorceIbCaba,@Rubro,@NroInsc,@FechaNroInsc,@CategoriaI,@CategoriaII, @FechaCategoria,
					@IbCiudadII,@Cai,@VtoCai,@Iso,@VtoIso,@Estado,@Califica,@FechaCalifica,@ObservacionesII,
					@Cufe,@CufeII,@CufeIII,@DirCufe,@DirCufeII,@DirCufeIII, @OrdFechaCalifica, @OrdFechaCategoria,
					@OrdFechaNroInsc, @Wdate)

		ELSE
			UPDATE surfactanSA.dbo.Proveedor	
			SET Nombre = @Nombre, Direccion = @Direccion, Localidad = @Localidad,Provincia = @Provincia, 
				Postal = @Postal,Region = @Region,Telefono = @Telefono,Dias = @Dias,
				Email = @Email,Observaciones = @Observaciones,Cuit = @Cuit, Tipo = @Tipo,Iva = @Iva,
				Cuenta = @Cuenta,NombreCheque = @NombreCheque,CodIb= @CodIb,CodIbCaba = @CodIbCaba,
				NroIb = @NroIb,	PorceIb=@PorceIb,PorceIbCaba=@PorceIbCaba,TipoProv = @Rubro,NroInsc=@NroInsc,
				FechaNroInsc=@FechaNroInsc,CategoriaI=@CategoriaI,CategoriaII=@CategoriaII, FechaCategoria=@FechaCategoria,
				IbCiudadII=@IbCiudadII,Cai=@Cai,VtoCai=@VtoCai,Iso=@Iso,VtoIso=@VtoIso,Estado=@Estado,Califica=@Califica,
				FechaCalifica=@FechaCalifica,ObservacionesII=@ObservacionesII,Cufe=@Cufe,CufeII=@CufeII,CufeIII=@CufeIII,
				DirCufe=@DirCufe,DirCufeII=@DirCufeII,DirCufeIII=@DirCufeIII, OrdFechaCalifica=@OrdFechaCalifica,
				OrdFechaCategoria=@OrdFechaCategoria,OrdFechaNroInsc=@OrdFechaNroInsc, Wdate=@Wdate
			WHERE Proveedor = @Proveedor
	ROLLBACK TRANSACTION
GO

CREATE PROCEDURE [dbo].[PR_baja_proveedor]
	(@proveedor varchar(11))
AS
	DECLARE @mensaje_error varchar(255)
	set @mensaje_error = ''
	BEGIN TRANSACTION
		BEGIN TRY
			DELETE FROM surfactanSA.dbo.Proveedor
			WHERE Proveedor = @proveedor
			IF @@ERROR = 0 COMMIT TRANSACTION
		END TRY
	BEGIN CATCH	
		ROLLBACK TRANSACTION
		IF (@mensaje_error = '') 
				set @mensaje_error = @mensaje_error + 'NO SE PUDO ELIMINAR EL PROVEEDOR.'	
		RAISERROR(@mensaje_error, 16, 217)
			WITH SETERROR
	END CATCH
GO

CREATE PROCEDURE [dbo].[PR_buscar_proveedor_por_codigo]
	(@codigo varchar(11))
AS
	SELECT Nombre, Direccion, Localidad,Provincia, Postal,Region,Telefono,Dias,
		Email,Observaciones,Cuit,Tipo,Iva,Cuenta,NombreCheque,CodIb,CodIbCaba,NroIb,
		PorceIb,PorceIbCaba,TipoProv,NroInsc,FechaNroInsc,CategoriaI,CategoriaII, FechaCategoria,
		IbCiudadII,Cai,VtoCai,Iso,VtoIso,Estado,Califica,FechaCalifica,ObservacionesII,
		Cufe,CufeII,CufeIII,DirCufe,DirCufeII,DirCufeIII
	FROM Proveedor p 
	WHERE p.Proveedor = @codigo

GO

CREATE PROCEDURE [dbo].[PR_buscar_proveedor_por_nombre]
	(@nombre varchar(50))
AS
	select LTRIM(RTRIM(p.Proveedor)) as Codigo, LTRIM(RTRIM(p.Nombre)) as Nombre
	from surfactanSA.dbo.Proveedor p
	where p.Nombre like '%' + @nombre + '%'
	order by p.Nombre

GO
