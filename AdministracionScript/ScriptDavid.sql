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
