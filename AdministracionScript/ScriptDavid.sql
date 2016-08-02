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
DROP PROCEDURE [dbo].[PR_alta_impCtaCtePrvNet]
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
	[Orden] [char](1) NULL,
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

CREATE PROCEDURE PR_alta_	
	(@clave varchar(26),
	@proveedor varchar(11),
	@letra varchar(1),
	@tipo varchar(2),
	@punto varchar(4),
	@numero varchar(4))
AS
	INSERT INTO dbo.impCtaCtePrvNet 
		(Clave, Proveedor, Letra, Tipo, Punto, Numero)
		VALUES
		(@clave, @proveedor, @letra, @tipo, @punto, @numero)
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
	from surfactanSA.dbo.CtaCtePrv CtaCtePrv
	WHERE CtaCtePrv.Proveedor between @proveedordesde  and @proveedorhasta 
		AND ((CtaCtePrv.Saldo <> 0 and @tipo = 'P')
			OR (@tipo = 'T')) 
	order by CtaCtePrv.Proveedor, CtaCtePrv.OrdFecha, CtaCtePrv.Tipo,CtaCtePrv.Numero


GO
