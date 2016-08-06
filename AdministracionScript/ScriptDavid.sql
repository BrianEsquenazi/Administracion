USE [surfactanSA]
GO

/*
----------------------------------------------------------------------------
										NOVEDADES
----------------------------------------------------------------------------
*/

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_corriente_proveedores_deuda]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cuenta_corriente_proveedores_deuda]
GO

USE [surfactanSA]
GO


SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[PR_buscar_cuenta_corriente_proveedores_deuda]
	(@proveedor varchar(11)
	, @tipo char(1))
AS
	select LTRIM(RTRIM(CtaCtePrv.Tipo)) as Tipo 
		 , LTRIM(RTRIM(CtaCtePrv.Letra)) as Letra
		 , LTRIM(RTRIM(CtaCtePrv.Punto)) as Punto
		 , LTRIM(RTRIM(CtaCtePrv.Numero)) as Numero
		 , CtaCtePrv.Total as Total
		 , CtaCtePrv.Saldo as Saldo
		 , LTRIM(RTRIM(CtaCtePrv.fecha)) as Fecha
		 , LTRIM(RTRIM(CtaCtePrv.Vencimiento)) as Vencimiento
		 
	from surfactanSA.dbo.CtaCtePrv CtaCtePrv
	WHERE CtaCtePrv.Proveedor = @proveedor 
		AND ((CtaCtePrv.Saldo <> 0 and @tipo = 'P')
			OR (@tipo = 'T')) 
	order by CtaCtePrv.Proveedor, CtaCtePrv.OrdFecha, CtaCtePrv.Tipo,CtaCtePrv.Numero

GO


/*
----------------------------------------------------------------------------
										PROCESOS
----------------------------------------------------------------------------
*/


/*
----------------------------------------------------------------------------
										ABM
----------------------------------------------------------------------------
*/

/*
----------------------------------------------------------------------------
									impCtaCtePrvNet
----------------------------------------------------------------------------
*/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[impCtaCtePrvNet]') AND type in (N'U'))
DROP TABLE [dbo].[impCtaCtePrvNet]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_limpiar_impCtaCtePrvNet]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_limpiar_impCtaCtePrvNet]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_impCtaCtePrvNet]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_impCtaCtePrvNet]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cuenta_corriente_proveedores_desdehasta]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cuenta_corriente_proveedores_desdehasta]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_actualizar_cuenta_corriente_proveedor]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_actualizar_cuenta_corriente_proveedor]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[impCtaCtePrvNet](
	[Clave] [char](26) NULL,
	[Proveedor] [char](11) NOT NULL,
	[Letra] [char](1) NULL,
	[Tipo] [char](2) NOT NULL,
	[Punto] [char](4) NULL,
	[Numero] [char](8) NOT NULL,
	[fecha] [char](10) NULL,
	[Estado] [char](1) NULL,
	[Vencimiento] [char](10) NULL,
	[Vencimiento1] [char](50) NULL,
	[Total] [float] NULL,
	[Saldo] [float] NULL,
	[OrdFecha] [char](8) NULL,
	[OrdVencimiento] [char](8) NULL,
	[Impre] [char](2) NULL,
	[Empresa] [smallint] NULL,
	[SaldoList] [float] NULL,
	[NroInterno] [int] NULL,
	[Orden] [int] NULL,
	[Acumulado] [float] NULL,
	[Titulo] [char](50) NULL,
	[Titulo1] [char](10) NULL,
	[Titulo2] [char](10) NULL,
	[Titulo3] [char](10) NULL,
	[Titulo4] [char](10) NULL,
	[Impre1] [float] NULL,
	[Impre2] [float] NULL,
	[Impre3] [float] NULL,
	[Impre4] [float] NULL,
	[Impre5] [float] NULL
	
) ON [PRIMARY]
GO

CREATE PROCEDURE PR_limpiar_impCtaCtePrvNet
AS
	DELETE 
	FROM dbo.impCtaCtePrvNet 
GO

CREATE PROCEDURE PR_alta_impCtaCtePrvNet	
	(@clave char(26),
	@proveedor char(11),
	@tipo char(2),
	@letra char(1),
	@punto char(4),
	@numero char(8),
	@total float,
	@saldo float,
	@fecha char(10),
	@vencimiento char(10),
	@vencimiento1 char(10),
	@impre char(2),
	@nrointerno int,
	@titulo char(50),
	@Acumulado float,
	@Orden float,
	@titulo1 char(10),
	@titulo2 char(10),
	@titulo3 char(10),
	@titulo4 char(10),
	@Impre1 float,
	@impre2 float,
	@impre3 float,
	@impre4 float,
	@Impre5 float)
AS
	INSERT INTO dbo.impCtaCtePrvNet 
		(Clave, Proveedor, Letra, Tipo, Punto, Numero, Total, Saldo, fecha, Vencimiento, Vencimiento1, Impre, NroInterno, Titulo, Acumulado, Orden, Titulo1, Titulo2, Titulo3,Titulo4,Impre1,Impre2,Impre3,Impre4,Impre5)
		VALUES
		(@clave, @proveedor, @letra, @tipo, @punto, @numero, @total, @saldo,@fecha, @vencimiento, @vencimiento1, @impre, @nrointerno, @titulo, @Acumulado, @orden, @titulo1, @titulo2, @titulo3, @titulo4, @Impre1, @impre2, @impre3, @impre4, @Impre5)
GO

CREATE PROCEDURE [dbo].[PR_buscar_cuenta_corriente_proveedores_desdehasta]
	(@proveedordesde char(11)
	,@proveedorhasta char(11)
	, @tipo char(1))
AS
	select LTRIM(RTRIM(CtaCtePrv.Tipo)) as Tipo 
		 , LTRIM(RTRIM(CtaCtePrv.Letra)) as Letra
		 , LTRIM(RTRIM(CtaCtePrv.Punto)) as Punto
		 , LTRIM(RTRIM(CtaCtePrv.Numero)) as Numero
		 , CtaCtePrv.Total as Total
		 , CtaCtePrv.Saldo as Saldo
		 , LTRIM(RTRIM(CtaCtePrv.fecha)) as Fecha
		 , LTRIM(RTRIM(CtaCtePrv.Vencimiento)) as Vencimiento
		 , LTRIM(RTRIM(CtaCtePrv.Vencimiento1)) as Vencimiento1
		 , LTRIM(RTRIM(CtaCtePrv.Impre)) as Impre
		 , CtaCtePrv.NroInterno as NroInterno
		 , LTRIM(RTRIM(CtaCtePrv.Clave)) as Clave
		 , LTRIM(RTRIM(CtaCtePrv.Proveedor)) as Proveedor

	from surfactanSA.dbo.CtaCtePrv CtaCtePrv
	WHERE CtaCtePrv.Proveedor between @proveedordesde  and @proveedorhasta 
		AND ((CtaCtePrv.Saldo <> 0 and @tipo = 'P')
			OR (@tipo = 'T')) 
	order by CtaCtePrv.Proveedor, CtaCtePrv.OrdFecha, CtaCtePrv.Tipo,CtaCtePrv.Numero


GO

CREATE PROCEDURE PR_actualizar_cuenta_corriente_proveedor
	@Tipo varchar(2)
	, @Letra varchar(1)
	, @Punto varchar(4)
	, @Numero varchar(8)
	, @Fecha varchar(10)
	, @Aplica float
	, @Proveedor varchar(11)
AS
BEGIN
	DECLARE @arreglo int = (-1) 
	IF (@Tipo = '03' or @Tipo = '05') 
		SET @arreglo = 1
	/*
	Lo pongo de esta forma porque el saldo esta almacenado como positivo para los primeros casos
	y negativo para los que toma el if por lo que para restar en un caso se debe restar y en el
	otro sumar
	*/ 

	UPDATE CtaCtePrv
	SET Saldo = Saldo + (@arreglo * @Aplica)
	WHERE Tipo = @Tipo
		and Letra = @Letra
		and Punto = @Punto
		and Numero = @Numero
		and fecha = @Fecha
		and Proveedor = @Proveedor
END
GO