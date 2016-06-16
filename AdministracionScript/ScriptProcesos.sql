USE [surfactanSA]
GO

/*
									ELIMINO PROCEDIMIENTOS 
*/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesogetcierre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesogetcierre]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_cierre]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_cierre]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesogetivacomp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesogetivacomp]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesogetivacompadicional]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesogetivacompadicional]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesogetctacte]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesogetctacte]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_depurar_cuentas_corrientes]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_depurar_cuentas_corrientes]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesoReteIb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesoReteIb]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesoReteGanan]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesoReteGanan]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesoReteIva]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesoReteIva]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesoReteIbRecibos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesoReteIbRecibos]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesoPerceIb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesoPerceIb]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_procesoPerceIva]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_procesoPerceIva]
GO

/*
									CREO PROCEDIMIENTOS 
*/

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

CREATE PROCEDURE PR_procesogetcierre (@mes int, @anio int)
AS
BEGIN
-- EN CASO DE SER AMBOS 0 INDICA QUE QUIERE VER TODOS LOS CIERRES
	IF (@mes = 0 AND @anio = 0)
		SELECT * FROM surfactanSA.dbo.Cierre
	ELSE
		SELECT [Mes]
			,[Ano]
			,[Estado]
		FROM surfactanSA.dbo.Cierre ci
		WHERE   ci.Ano = @anio and @mes = ci.Mes 
END
GO

CREATE PROCEDURE PR_alta_cierre (@mes int, @anio int, @estado int)
AS
BEGIN
-- DE EXISTIR YA EL CIERRE INDICADO, LO MODIFICA, DE OTRA FORMA LO INSERTA
	IF(EXISTS (SELECT 1 
				FROM surfactanSA.dbo.Cierre ci 
				WHERE ci.Ano = @anio and ci.Mes = @mes))
		UPDATE surfactanSA.dbo.Cierre
			SET Estado = @estado
			WHERE Ano = @anio 
				and Mes = @mes
	ELSE
		INSERT INTO surfactanSA.dbo.Cierre
			values (@mes, @anio, @estado)
	
END
GO

CREATE PROCEDURE PR_procesogetivacomp (@desde varchar(10)
								, @hasta varchar(10)
								, @letra varchar(1)
								, @tipo varchar(2))
AS
BEGIN

--REVISAR PORQUE NO FUNCIONA AL COMPARAR PERIODOS ORDENABLES

	declare @desde_fecha varchar(10) = (select dbo.[FN_get_fecha_desde_ordenable] (@desde) )
	declare @hasta_fecha varchar(10) = (select dbo.[FN_get_fecha_desde_ordenable] (@hasta) )

	SELECT ic.[NroInterno]
      ,ic.[Proveedor]
      ,ic.[Tipo]
      ,ic.[Letra]
      ,ic.[Punto]
      ,ic.[Numero]
      ,ic.[Fecha]
      ,ic.[Periodo]
      ,ic.[Neto]
      ,ic.[Iva21]
      ,ic.[Iva5]
      ,ic.[Iva27]
      ,ic.[Ib]
      ,ic.[Exento]
      ,ic.[Despacho]
      ,ic.[Iva105]
      ,p.Nombre
      ,p.Cuit
  FROM [surfactanSA].[dbo].[IvaComp] ic
  JOIN Proveedor p on p.Proveedor = ic.Proveedor
  WHERE --dbo.FN_procesogetfecha_ordenable (ic.Periodo) between @desde and  @hasta
	ic.Periodo between @desde_fecha and @hasta_fecha
  	and ic.Letra = @letra
	and ic.Tipo = @tipo
END
GO

CREATE PROCEDURE PR_procesogetivacompadicional (@clave varchar(10))
AS
BEGIN
	SELECT
		ica.Clave
		, ica.Tipo
		, ica.Letra
		, ica.Punto
		, ica.Numero
		, ica.Fecha
		, ica.Neto
		, ica.Iva21
		, ica.PerceIb
		, ica.Iva27
		, ica.PerceIva
		, ica.Iva105
		, ica.Exento
	FROM surfactanSA.dbo.IvaCompAdicional ica
	WHERE ica.Clave = @clave
END
GO

CREATE PROCEDURE PR_procesogetctacte (@desde varchar(8), @hasta varchar(8))
AS
BEGIN
	SELECT 
		cu.OrdFecha
		, cu.Tipo
		, cu.Numero
		, cu.fecha
		, cu.Cliente
		, cu.Neto
		, cu.Iva1
		, cu.Iva2
		, cu.ImpoIbTucu
		, cu.ImpoIbCiudad
		, cu.ImpoIb
		, cu.Vencimiento
		, cli.Razon
		, cli.Cuit
	FROM surfactanSA.dbo.CtaCte cu
	JOIN Cliente cli on cli.Cliente = cu.Cliente
	WHERE cu.OrdFecha between @desde and @hasta
END
GO

CREATE PROCEDURE PR_depurar_cuentas_corrientes 
AS
BEGIN
-- AGREGO SALDO <> 0 PARA QUE MODIFIQUE MENOS 
-- REGISTOS Y DE ESTA FORMA SEA MAS RAPIDO
	UPDATE surfactanSA.dbo.CtaCte
	SET Saldo = 0
	WHERE Saldo between (-0.1) and (0.1)
		and Saldo <> 0
	 
	UPDATE surfactanSA.dbo.CtaCtePrv
	SET Saldo = 0
	WHERE Saldo between (-0.1) and (0.1) 
		and Saldo <> 0
END
GO

CREATE PROCEDURE PR_procesoReteIb
	(@desde varchar(8)
	, @hasta varchar(8))
AS
BEGIN
	SELECT
		pa.FechaOrd
		, pa.RetOtra
		, pa.Renglon
		, pa.Proveedor
		, pa.Fecha
		, pa.Orden
		, pr.Cuit
	FROM surfactanSA.dbo.Pagos pa
	JOIN Proveedor pr on pr.Proveedor = pa.Proveedor
	WHERE pa.FechaOrd between @desde and @hasta
		and pa.RetOtra <> 0
		and pa.Renglon = 1
END
GO

CREATE PROCEDURE PR_procesoReteGanan
	(@desde varchar(8)
	, @hasta varchar(8))
AS
BEGIN
	SELECT
		pa.FechaOrd
		, pa.Orden
		, pa.Renglon
		, pa.Importe
		, pa.Fecha
		, pa.Retencion
		, pa.Proveedor
		, pa.CertificadoGan
		, pr.Cuit
		, pr.Tipo
	FROM surfactanSA.dbo.Pagos pa
	JOIN Proveedor pr on pr.Proveedor = pa.Proveedor
	WHERE pa.FechaOrd between @desde and @hasta
		and pa.RetOtra <> 0
		and pa.Renglon = 1
END
GO

CREATE PROCEDURE PR_procesoReteIva
	(@desde varchar(8)
	, @hasta varchar(8))
AS
BEGIN
	SELECT
		re.Fechaord
		, re.RetIva
		, re.Renglon
		, re.Cliente
		, re.Fecha
		, re.ComproIva
		, cli.Cuit
	FROM surfactanSA.dbo.Recibos re
	JOIN Cliente cli on cli.Cliente = re.Cliente
	WHERE re.Fechaord between @desde and @hasta
		and re.RetOtra <> 0
		and re.Renglon = 1
END
GO

CREATE PROCEDURE PR_procesoReteIbRecibos
	(@desde varchar(8)
	, @hasta varchar(8))
AS
BEGIN
	SELECT
		re.Fechaord
		, re.RetOtra
		, re.Renglon
		, re.Cliente
		, re.Fecha
		, re.ComproIb
		, cli.Cuit
	FROM surfactanSA.dbo.Recibos re
	JOIN Cliente cli on cli.Cliente = re.Cliente
	WHERE re.Fechaord between @desde and @hasta
		and re.RetOtra <> 0
		and re.Renglon = 1
END
GO

CREATE PROCEDURE PR_procesoPerceIb
	(@desde varchar(8)
	, @hasta varchar(8))
AS
BEGIN
	SELECT
		cc.OrdFecha
		, cc.ImpoIb
		, cc.Clave
		, cc.fecha
		, cc.Tipo
		, cc.Numero
		, cc.Cliente
		, cc.Neto	
		, cli.Cuit
	FROM surfactanSA.dbo.CtaCte cc
	JOIN Cliente cli on cli.Cliente = cc.Cliente
	WHERE cc.OrdFecha between @desde and @hasta
		and cc.ImpoIb <> 0

END
GO

CREATE PROCEDURE PR_procesoPerceIva
	(@desde varchar(8)
	, @hasta varchar(8))
AS
BEGIN

	declare @periodo_desde varchar(10)
	declare @periodo_hasta varchar(10)
	SET @periodo_desde = (select dbo.[FN_get_fecha_desde_ordenable] (@desde) )
	SET @periodo_hasta = (select dbo.[FN_get_fecha_desde_ordenable] (@hasta) )
	
	SELECT top 10
		ic.Periodo
		, ic.Iva5
		, ic.Punto
		, ic.Numero
		, ic.Proveedor
		, ic.Fecha
		, ic.Fecha
		, p.Nombre
		, p.Cuit
	FROM surfactanSA.dbo.IvaComp ic
	JOIN Proveedor p on p.Proveedor = ic.Proveedor
	WHERE ic.Periodo between @periodo_desde and @periodo_hasta
		and ic.Iva5 <> 0

END
GO