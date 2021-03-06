if exists (select * from sysobjects where id = object_id(N'[dbo].[AltaProveedor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[AltaProveedor]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[BorrarProveedor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[BorrarProveedor]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[ConsultaCuentas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ConsultaCuentas]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[ConsultaProveedores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ConsultaProveedores]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[Empresas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[Empresas]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[ListaCuentas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ListaCuentas]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[ListaProveedores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ListaProveedores]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[ModificaProveedor]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ModificaProveedor]
GO

if exists (select * from sysobjects where id = object_id(N'[dbo].[SP_ConsultaBancos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[SP_ConsultaBancos]
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE AltaProveedor
             @Proveedor char(11),
	@Nombre char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Postal char(2),
	@Cuit char(4),
	@Telefono    char(15),
	@EMail char(30),
	@Observaciones char(30),
	@Dias char(50),
	@Tipo char(1),
	@Iva char(1),
	@Provincia char(1),
	@Cuenta char(1),
	@NombreCheque char(1)

 AS
	INSERT INTO  Proveedor
			(
			Proveedor   ,
			Nombre        ,                                     
			Direccion      ,                                    
			Localidad      ,                                    
			Provincia ,
			Postal ,
			Cuit     ,       
			Telefono                       ,
			Email                          ,
			Observaciones            ,                          
			Tipo ,
			Iva  ,
			Dias ,                
			Cuenta     ,
			NombreCheque                                       
			)
VALUES
			(
		             @Proveedor,
			@Nombre,
			@Direccion,
			@Localidad,
			@Provincia,
			@Postal,
			@Cuit,
			@Telefono,
			@EMail,
			@Observaciones,
			@Tipo,
			@Iva,
			@Dias,
			@Cuenta,
			@NombreCheque
			)
GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE BorrarProveedor
	@Proveedor   char(11)
 AS

DELETE	Proveedor
WHERE
		Proveedor = @Proveedor
	
GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE ConsultaCuentas
	@Cuenta char(10)	
 AS

SELECT * FROM Cuenta
WHERE
	Cuenta = @Cuenta


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE ConsultaProveedores
	@Proveedor  char(11)
 AS

SELECT * FROM Proveedor
WHERE
	Proveedor = @Proveedor


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

Create procedure Empresas
as
select * from empresa

GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE ListaCuentas  AS

Select * From Cuenta
Order by Descripcion






GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE Procedure ListaProveedores AS

Select * From Proveedor
Order by Proveedor


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

CREATE PROCEDURE ModificaProveedor
            @Proveedor char(11),
	@Nombre char(50),
	@Direccion char(50),
	@Localidad char(50),
	@Postal char(2),
	@Cuit char(4),
	@Telefono    char(15),
	@EMail char(30),
	@Observaciones char(30),
	@Dias char(50),
	@Tipo char(1),
	@Iva char(1),
	@Provincia char(2),
	@Cuenta char(1),
	@NombreCheque char(1)

 AS

UPDATE  Proveedor
	SET
--		Proveedor   	= @Proveedor,
		Nombre       	= @Nombre,                                     
		Direccion      	= @Direccion,                                    
		Localidad      	= @Localidad,                                    
		Provincia 	= @Provincia,
		Postal 		= @Postal,
		Cuit    		= @Cuit,       
		Telefono	= @Telefono                      ,
		Email         	= @Email               ,
		Observaciones	= @Observaciones            ,                          
		Tipo 		= @Tipo,
		Iva  		= @Iva,
		Dias		= @Dias,                
		Cuenta 		= @Cuenta   ,
		NombreCheque 	= @NombreCheque                                       
WHERE
	Proveedor = @Proveedor


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO


CREATE PROCEDURE SP_ConsultaBancos
AS

SELECT * FROM Bancos


GO
SET QUOTED_IDENTIFIER  OFF    SET ANSI_NULLS  ON 
GO

