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
									IMPCTACTEPRV
----------------------------------------------------------------------------
*/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ImpCtaCtePrv]') AND type in (N'U'))
DROP TABLE [dbo].[ImpCtaCtePrv]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_limpiar_impCtaCtePrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_limpiar_impCtaCtePrv]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_impCtaCtePrv]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_impCtaCtePrv]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ImpCtaCtePrv](
	[Clave] [nvarchar](26) NULL,
	[Proveedor] [nvarchar](11) NOT NULL,
	[Letra] [nvarchar](1) NULL,
	[Tipo] [nvarchar](2) NOT NULL,
	[Punto] [nvarchar](4) NULL,
	[Numero] [nvarchar](8) NOT NULL,
	[fecha] [nvarchar](10) NULL,
	[Estado] [nvarchar](1) NULL,
	[Vencimiento] [nvarchar](10) NULL,
	[Vencimiento1] [nvarchar](50) NULL,
	[Total] [float] NULL,
	[Saldo] [float] NULL,
	[OrdFecha] [nvarchar](8) NULL,
	[OrdVencimiento] [nvarchar](8) NULL,
	[Impre] [nvarchar](2) NULL,
	[Empresa] [smallint] NULL,
	[SaldoList] [float] NULL,
	[NroInterno] [int] NULL,
	[Lista] [nvarchar](1) NULL,
	[Acumulado] [float] NULL
) ON [PRIMARY]
GO

CREATE PROCEDURE PR_limpiar_impCtaCtePrv
AS
	DELETE 
	FROM dbo.ImpCtaCtePrv 
GO

CREATE PROCEDURE PR_alta_impCtaCtePrv 
	(@clave varchar(26),
	@proveedor varchar(11),
	@letra varchar(1),
	@tipo varchar(2),
	@punto varchar(4),
	@numero varchar(4))
AS
	INSERT INTO dbo.ImpCtaCtePrv 
		(Clave, Proveedor, Letra, Tipo, Punto, Numero)
		VALUES
		(@clave, @proveedor, @letra, @tipo, @punto, @numero)
GO
