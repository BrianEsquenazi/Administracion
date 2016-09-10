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

-- AGREGA LAS COLUMNAS NUEVAS A CtaCtePrv en caso de que no existan
IF NOT EXISTS(
    SELECT *
    FROM sys.columns 
    WHERE Name      = N'TituloI'
      AND Object_ID = Object_ID(N'CtaCtePrv'))
BEGIN
	ALTER TABLE [CtaCtePrv]
    ADD [TituloI] char(50) NULL 
END

IF NOT EXISTS(
    SELECT *
    FROM sys.columns 
    WHERE Name      = N'TituloII'
      AND Object_ID = Object_ID(N'CtaCtePrv'))
BEGIN
	ALTER TABLE [CtaCtePrv]
    ADD [TituloII] char(50) NULL 
END

IF NOT EXISTS(
    SELECT *
    FROM sys.columns 
    WHERE Name      = N'Auxi1'
      AND Object_ID = Object_ID(N'CtaCtePrv'))
BEGIN
	ALTER TABLE [CtaCtePrv]
    ADD [Auxi1] char(10) NULL 
END

IF NOT EXISTS(
    SELECT *
    FROM sys.columns 
    WHERE Name      = N'Auxi1'
      AND Object_ID = Object_ID(N'CtaCtePrv'))
BEGIN
	ALTER TABLE [CtaCtePrv]
    ADD [Auxi1] char(10) NULL 
END

IF NOT EXISTS(
    SELECT *
    FROM sys.columns 
    WHERE Name      = N'Auxi3'
      AND Object_ID = Object_ID(N'CtaCtePrv'))
BEGIN
	ALTER TABLE [CtaCtePrv]
    ADD [Auxi3] char(8) NULL 
END

IF NOT EXISTS(
    SELECT *
    FROM sys.columns 
    WHERE Name      = N'Auxi4'
      AND Object_ID = Object_ID(N'CtaCtePrv'))
BEGIN
	ALTER TABLE [CtaCtePrv]
    ADD [Auxi4] char(8) NULL 
END


/*
----------------------------------------------------------------------------
										PROCESOS
----------------------------------------------------------------------------
*/


/*
----------------------------------------------------------------------------
									TABLAS MODIFICADAS
----------------------------------------------------------------------------
*/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Movban]') AND type in (N'U'))
DROP TABLE [dbo].[Movban]
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Movban](
	[da] [int] NULL,
	[Banco] [int] NULL,
	[Fecha] [char](10) NOT NULL,
	[FechaOrd] [char](8) NULL,
	[Acredita] [char](10) NOT NULL,
	[AcreditaOrd] [char](8) NULL,
	[Observaciones] [char](50) NULL,
	[Numero] [char](10) NULL,
	[Debito] [float] NULL,
	[Credito] [float] NULL,
	[Comprobante] [char](8) NULL,
	[Empresa] [int] NULL,
	[Titulo] [char](50) NULL,
	[Titulo1] [char](50) NULL,
	[Proveedor] [char](11) NOT NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO


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

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[impcyb]') AND type in (N'U'))
DROP TABLE [dbo].[impcyb]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_limpiar_impcyb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_limpiar_impcyb]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_impCyb]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_impCyb]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_pagos_fecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_pagos_fecha]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_recibos_fecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_recibos_fecha]
GO



IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_depositos_fecha]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_depositos_fecha]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ListaIvaCompras]') AND type in (N'U'))
DROP TABLE [dbo].[ListaIvaCompras]
GO


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_limpiar_ListaIvaCompras]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_limpiar_ListaIvaCompras]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_ListaIvaCompras]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_ListaIvaCompras]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_Lee_IvaComp]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_Lee_IvaComp]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_Lee_IvaCompAdicional]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_Lee_IvaCompAdicional]
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


CREATE TABLE [dbo].[impcyb](
	[Clave] [char](24) NULL,
	[TipoMovi] [char](1) NOT NULL,
	[Proveedor] [char](11) NULL,
	[TipoComp] [char](2) NOT NULL,
	[LetraComp] [char](1) NULL,
	[PuntoComp] [char](4) NOT NULL,
	[NroComp] [char](8) NULL,
	[Renglon] [char](2) NULL,
	[Fecha] [char](10) NULL,
	[Observaciones] [char](50) NULL,
	[Cuenta] [char](10) NULL,
	[Debito] [float] NULL,
	[Credito] [float] NULL,
	[FechaOrd] [char](8) NULL,
	[Titulo] [char](50) NULL,
	[Empresa] [smallint] NULL,
	[DebitoList] [float] NULL,
	[CreditoList] [float] NULL,
	[NroINterno] [int] NULL,
	[Nombre] [char](50) NULL,
	[TituloList] [char](50) NULL,
	[Varios] [char](50) NULL,
	[ClaveOrd] [char](20) NULL
	
) ON [PRIMARY]
GO


CREATE PROCEDURE PR_limpiar_impcyb
AS
	DELETE 
	FROM dbo.impcyb
GO


CREATE PROCEDURE PR_alta_impcyb
	(@clave char(24),
	@TipoMovi char(1),
	@NroINterno int,
	@Proveedor char(11),
	@TipoComp char(2),
	@LetraComp char(1),
	@PuntoComp char(4),
	@NroComp char(8),
	@Renglon char(2),
	@Fecha char(10),
	@Observaciones char(50),
	@Cuenta char(10),
	@Credito float,
	@Debito float,
	@FechaOrd char(8),
	@Titulo char(50),
	@Empresa smallint,
	@TituloList char(50),
	@Varios char(50),
	@ClaveOrd char(20))
AS
	INSERT INTO dbo.impcyb
		(Clave, TipoMovi, NroInterno, Proveedor, TipoComp, LetraComp, PuntoComp, NroComp, Renglon, Fecha, Observaciones, Cuenta, Credito, 
		Debito, FechaOrd, Titulo, Empresa, TituloList, Varios, ClaveOrd)
		VALUES
		(@clave, @TipoMovi, @NroInterno, @Proveedor, @TipoComp, @LetraComp, @PuntoComp, @NroComp, @Renglon, @Fecha, @Observaciones, @Cuenta, @Credito, 
		@Debito, @FechaOrd, @Titulo, @Empresa, @TituloList, @Varios, @ClaveOrd)
GO


CREATE PROCEDURE [dbo].[PR_buscar_pagos_fecha]
	(@DesdeFecha char(8)
	, @HastaFecha char(8))
AS
	select Pagos.FechaOrd as FechaOrd
		 , Pagos.Orden as Orden
		 , Pagos.TipoOrd as TipoOrd
		 , Pagos.Banco2 as Banco2
		 , Pagos.Cuenta as Cuenta
		 , Pagos.Proveedor as Proveedor
		 , Pagos.Letra1 as Letra1
		 , Pagos.Tipo1 as Tipo1
		 , Pagos.Punto1 as Punto1
		 , Pagos.Numero1 as Numero1
		 , Pagos.Importe1 as Importe1
		 , Pagos.Fecha as Fecha
		 , Pagos.TipoReg as Tiporeg
		 , Pagos.Observaciones as Observaciones
		 , Pagos.RetOtra as RetOtra
		 , Pagos.Retencion as Retencion
		 , Pagos.RetIbCiudad as RetIbCiudad
		 , Pagos.Importe2 as Importe2
		 , Pagos.Tipo2 as Tipo2
		 , Pagos.Numero2 as Numero2
		 , Pagos.Renglon as Renglon
		 , Prove.Provincia as Provincia
		 
	from surfactanSA.dbo.pagos pagos
	JOIN Proveedor Prove on Prove.Proveedor = Pagos.Proveedor
	WHERE pagos.FechaOrd between @DesdeFecha and @HastaFecha
	order by pagos.Clave



GO





CREATE PROCEDURE [dbo].[PR_buscar_recibos_fecha]
	(@DesdeFecha char(8)
	, @HastaFecha char(8))
AS
	select Recibos.FechaOrd as FechaOrd
		 , recibos.Recibo as Recibo
		 , recibos.TipoReg as TipoReg
		 , recibos.TipoRec as TipoRec
		 , recibos.Cuenta as Cuenta
		 , recibos.Cliente as Cliente
		 , recibos.Tipo1 as Tipo1
		 , recibos.letra1 as Letra1
		 , recibos.Punto1 as Punto1
		 , recibos.Numero1 as Numero1
		 , recibos.Fecha as Fecha
		 , recibos.Tipo2 as Tipo2
		 , recibos.Numero2 as Numero2
		 , recibos.Importe1 as Importe1
		 , recibos.Paridad as Paridad
		 , recibos.Importe2 as Importe2
		 , recibos.RetIva as RetIva
		 , recibos.RetOtra as RetOtra
		 , recibos.RetSuss as RetSuss
		 , recibos.RetGanancias as RetGanancias
		 , recibos.Renglon as Renglon
		 , Clie.Provincia as Provincia
		 
	from surfactanSA.dbo.recibos recibos
	JOIN Cliente Clie on Clie.Cliente = recibos.Cliente
	WHERE recibos.FechaOrd between @DesdeFecha and @HastaFecha
	order by recibos.Clave

GO








CREATE PROCEDURE [dbo].[PR_buscar_depositos_fecha]
	(@DesdeFecha char(8)
	, @HastaFecha char(8))
AS
	select Depositos.FechaOrd as FechaOrd
		 , Depositos.Deposito as Deposito
		 , Depositos.Tipo2 as Tipo2
		 , Depositos.Numero2 as Numero2
		 , Depositos.Fecha as Fecha
		 , Depositos.Importe2 as Importe2
		 , Depositos.Banco as Banco
		 
	from surfactanSA.dbo.Depositos depositos
	WHERE depositos.FechaOrd between @DesdeFecha and @HastaFecha
	order by depositos.Clave

GO









--
--   ---------------------------------------------------------------------------------------------------------
--


CREATE TABLE [dbo].[ListaIvaCompras](
	[NroInterno] [int] NULL,
	[Proveedor] [char](11) NULL,
	[Tipo] [char](2) NOT NULL,
	[Letra] [char](1) NULL,
	[Punto] [char](4) NOT NULL,
	[Numero] [char](8) NULL,
	[Fecha] [char](10) NULL,
	[Periodo] [char](10) NULL,
	[Neto] [float] NULL,
	[Iva21] [float] NULL,
	[Iva5] [float] NULL,
	[Iva27] [float] NULL,
	[Iva105] [float] NULL,
	[Ib] [float] NULL,
	[Exento] [float] NULL,
	[Impre] [char](3) NOT NULL,
	[OrdFecha] [char](8) NOT NULL,
	[Titulo] [char](50) NOT NULL,
	[TituloII] [char](50) NOT NULL,
	[Nombre] [char](50) NOT NULL,
	[Cuit] [char](15) NOT NULL
	
) ON [PRIMARY]
GO



--
--   ---------------------------------------------------------------------------------------------------------
--

CREATE PROCEDURE PR_limpiar_ListaIvaCompras
AS
	DELETE 
	FROM dbo.ListaIvaCompras
GO



--
--   ---------------------------------------------------------------------------------------------------------
--

CREATE PROCEDURE PR_alta_ListaIvaCompras
	(@NroINterno int,
	@Proveedor char(11),
	@Tipo char(2),
	@Letra char(1),
	@Punto char(4),
	@Numero char(8),
	@Fecha char(10),
	@Periodo char(10),
	@Neto float,
	@Iva21 float,
	@Iva5 float,
	@Iva27 float,
	@Iva105 float,
	@Ib float,
	@Exento float,
	@Impre char(3),
	@OrdFecha char(8),
	@Titulo char(50),
	@TituloII char(50),
	@Nombre char(50),
	@Cuit char(15))
AS
	INSERT INTO dbo.ListaIvaCompras
		(NroINterno, Proveedor, tipo, Letra, Punto, Numero, Fecha, Periodo, Neto, Iva21, Iva5, Iva27, Iva105, Ib, 
		 Exento, Impre, OrdFecha, Titulo, TituloII, Nombre, cuit)
		VALUES
		(@NroINterno, @Proveedor, @tipo, @Letra, @Punto, @Numero, @Fecha, @Periodo, @Neto, @Iva21, @Iva5, @Iva27, @Iva105, @Ib, 
		 @Exento, @Impre, @OrdFecha, @Titulo, @TituloII, @Nombre, @Cuit)
GO





--
--   ---------------------------------------------------------------------------------------------------------
--


CREATE PROCEDURE [dbo].[PR_Lee_IvaComp] (@desde_fecha varchar(8)
								, @hasta_Fecha varchar(8))
AS
BEGIN

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
      ,ISNULL(ic.[Iva105],0)
      ,p.Nombre
      ,p.Cuit
  FROM [surfactanSA].[dbo].[IvaComp] ic
  JOIN Proveedor p on p.Proveedor = ic.Proveedor
  WHERE dbo.FN_get_fecha_ordenable (ic.Periodo) between @desde_fecha and  @hasta_fecha
END

GO






--
--   ---------------------------------------------------------------------------------------------------------
--


CREATE PROCEDURE [dbo].[PR_Lee_IvaCompAdicional] (@nrointerno integer)
AS
BEGIN

	SELECT ic.[NroInterno]
      ,ic.[Tipo]
      ,ic.[Letra]
      ,ic.[Punto]
      ,ic.[Numero]
      ,ic.[Fecha]
      ,ic.[Neto]
      ,ic.[Iva21]
      ,ic.[perceiva]
      ,ic.[Iva27]
      ,ic.[perceIb]
      ,ic.[Exento]
      ,ISNULL(ic.[Iva105],0)
      ,ic.[Cuit]
      ,ic.[Razon]
  FROM [surfactanSA].[dbo].[IvaCompAdicional] ic
  WHERE ic.NroInterno = @nrointerno
END

GO




--
--   ---------------------------------------------------------------------------------------------------------
--


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_limpiar_MovBan]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_limpiar_MovBan]
GO
CREATE PROCEDURE PR_limpiar_MovBan
AS
	DELETE 
	FROM dbo.Movban
GO






--
--   ---------------------------------------------------------------------------------------------------------
--


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_pagos_movban]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_pagos_movban]
GO


CREATE PROCEDURE [dbo].[PR_buscar_pagos_movban]
	(@DesdeFecha char(8)
	, @HastaFecha char(8)
	, @DesdeBanco int
	, @HastaBanco int)
AS
	select Pagos.FechaOrd as FechaOrd
		 , Pagos.Orden as Orden
		 , Pagos.TipoOrd as TipoOrd
		 , Pagos.Banco2 as Banco2
		 , Pagos.Cuenta as Cuenta
		 , Pagos.Proveedor as Proveedor
		 , Pagos.Fecha as Fecha
		 , Pagos.Fecha2 as Fecha2
		 , Pagos.FechaOrd2 as FechaOrd2
		 , Pagos.TipoReg as Tiporeg
		 , Pagos.Observaciones as Observaciones
		 , Pagos.Importe2 as Importe2
		 , Pagos.Tipo2 as Tipo2
		 , Pagos.Numero2 as Numero2
		 , Pagos.Renglon as Renglon
		 , Pagos.Importe1 as Importe1
		 
	from surfactanSA.dbo.pagos pagos
	JOIN Proveedor Prove on Prove.Proveedor = Pagos.Proveedor
	WHERE pagos.FechaOrd between @DesdeFecha and @HastaFecha
		and pagos.banco2 between @DesdeBanco and @HastaBanco
		and pagos.Tipo2 = '02'
	order by pagos.Clave

GO




--
--   ---------------------------------------------------------------------------------------------------------
--

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_pagos_movbanII]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_pagos_movbanII]
GO


CREATE PROCEDURE [dbo].[PR_buscar_pagos_movbanII]
	(@DesdeFecha char(8)
	, @HastaFecha char(8)
	, @DesdeBanco int
	, @HastaBanco int)
AS
	select Pagos.FechaOrd as FechaOrd
		 , Pagos.Orden as Orden
		 , Pagos.TipoOrd as TipoOrd
		 , Pagos.Banco2 as Banco2
		 , Pagos.Cuenta as Cuenta
		 , Pagos.Proveedor as Proveedor
		 , Pagos.Fecha as Fecha
		 , Pagos.Fecha2 as Fecha2
		 , Pagos.FechaOrd2 as FechaOrd2
		 , Pagos.TipoReg as Tiporeg
		 , Pagos.Observaciones as Observaciones
		 , Pagos.Importe2 as Importe2
		 , Pagos.Tipo2 as Tipo2
		 , Pagos.Numero2 as Numero2
		 , Pagos.Renglon as Renglon
		 , Pagos.Importe1 as Importe1
		 
	from surfactanSA.dbo.pagos pagos
	--JOIN Proveedor Prove on Prove.Proveedor = Pagos.Proveedor
	WHERE pagos.FechaOrd between @DesdeFecha and @HastaFecha
		and pagos.banco2 between @DesdeBanco and @HastaBanco
		and pagos.banco2 <> 0
		--and pagos.TipoReg = '1'
	order by pagos.Clave

GO










--
--   ---------------------------------------------------------------------------------------------------------
--


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_depositos_movban]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_depositos_movban]
GO

CREATE PROCEDURE [dbo].[PR_buscar_depositos_movban]
	(@DesdeFecha char(8)
	, @HastaFecha char(8)
	, @DesdeBanco int
	, @HastaBanco int)
AS
	select Depositos.FechaOrd as FechaOrd
		 , Depositos.Deposito as Deposito
		 , Depositos.Tipo2 as Tipo2
		 , Depositos.Numero2 as Numero2
		 , Depositos.Fecha as Fecha
		 , Depositos.Importe2 as Importe2
		 , Depositos.Banco as Banco
		 , Depositos.Importe as Importe
		 , Depositos.Acredita as Acredita
		 , Depositos.AcreditaOrd as AcreditaOrd
		 
	from surfactanSA.dbo.Depositos depositos
	WHERE depositos.FechaOrd between @DesdeFecha and @HastaFecha
		and depositos.banco between @DesdeBanco and @HastaBanco
		and depositos.Renglon = 1
	order by depositos.Clave


GO








--
--   ---------------------------------------------------------------------------------------------------------
--

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_recibos_movban]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_recibos_movban]
GO

CREATE PROCEDURE [dbo].[PR_buscar_recibos_movban]
	(@DesdeFecha char(8)
	, @HastaFecha char(8))
AS
	select Recibos.FechaOrd as FechaOrd
		 , recibos.Recibo as Recibo
		 , recibos.TipoReg as TipoReg
		 , recibos.TipoRec as TipoRec
		 , recibos.Cuenta as Cuenta
		 , recibos.Cliente as Cliente
		 , recibos.Tipo1 as Tipo1
		 , recibos.letra1 as Letra1
		 , recibos.Punto1 as Punto1
		 , recibos.Numero1 as Numero1
		 , recibos.Fecha as Fecha
		 , recibos.Tipo2 as Tipo2
		 , recibos.Numero2 as Numero2
		 , recibos.Importe1 as Importe1
		 , recibos.Paridad as Paridad
		 , recibos.Importe2 as Importe2
		 , recibos.RetIva as RetIva
		 , recibos.RetOtra as RetOtra
		 , recibos.RetSuss as RetSuss
		 , recibos.RetGanancias as RetGanancias
		 , recibos.Renglon as Renglon
		 , Clie.Provincia as Provincia
		 
	from surfactanSA.dbo.Recibos recibos
	JOIN Cliente Clie on Clie.Cliente = recibos.Cliente
	WHERE recibos.Fechaord between @DesdeFecha and @HastaFecha
		and (recibos.Cuenta = '21' or recibos.Cuenta = '22' or recibos.Cuenta = '25' or 
		     recibos.Cuenta = '26' or recibos.Cuenta = '27')
	order by recibos.Clave


GO




--
--   ---------------------------------------------------------------------------------------------------------
--


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_saldo_inicial_pagos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_saldo_inicial_pagos]
GO

CREATE PROCEDURE PR_get_saldo_inicial_pagos
	@desde varchar(8)
	, @hasta varchar(8)
	, @desde_banco smallint
	, @hasta_banco smallint
AS
BEGIN
	DECLARE @importe float
-- El isnull interno es para que valgan 0 todos los nullos que pasen el filtro
-- el isnull externo es para que, en caso de que nada pase el filtro, no retorne null sino 0
	SET @importe = (SELECT ISNULL( SUM( ISNULL(p.Importe2,0) ) ,0) 
	FROM Pagos p
	WHERE
		p.Tipo2 = '02'
		AND p.banco2 BETWEEN @desde_banco AND @hasta_banco
		AND p.FechaOrd2 BETWEEN @desde AND @hasta ) 
		
	RETURN @importe
END
GO





--
--   ---------------------------------------------------------------------------------------------------------
--


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_MovBan]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_MovBan]
GO


CREATE PROCEDURE PR_alta_MovBan
	(@da int,
	@Banco int,
	@Fecha char(10),
	@FechaOrd char(8),
	@Acredita char(10),
	@AcreditaOrd char(8),
	@Observaciones char(50),
	@Numero char(10),
	@Debito float,
	@Credito float,
	@Comprobante char(8),
	@Empresa int,
	@Titulo char(50),
	@Titulo1 char(50),
	@Proveedor char(11))
AS
	INSERT INTO dbo.Movban
		(da, Banco, Fecha, FechaOrd, Acredita, AcreditaOrd, Observaciones, Numero, Debito, Credito, Comprobante, Empresa, Titulo, Titulo1, 
		 Proveedor)
		VALUES
		(@da, @Banco, @Fecha, @FechaOrd, @Acredita, @AcreditaOrd, @Observaciones, @Numero, @Debito, @Credito, @Comprobante, @Empresa, @Titulo, @Titulo1, 
		 @Proveedor)
GO



--
--   ---------------------------------------------------------------------------------------------------------
--

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_lee_trabajo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_lee_trabajo]
GO


USE [surfactanSA]
GO

/****** Object:  StoredProcedure [dbo].[PR_lee_trabajo]    Script Date: 08/25/2016 11:22:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[PR_lee_trabajo]
	(@codigo integer)
AS
	select Trabajo.Codigo as Codigo
		 , Trabajo.Proceso as Proceso
		 , Trabajo.Destino as Destino 
		 , Trabajo.Planta as Planta
		 , Trabajo.Orden as Orden 
		 , Trabajo.Nombre as Nombre
	from surfactanSA.dbo.Trabajo Trabajo
	WHERE Trabajo.Codigo = @Codigo



GO



--
--   ---------------------------------------------------------------------------------------------------------
--

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Valcar]') AND type in (N'U'))
DROP TABLE [dbo].[Valcar]
GO


CREATE TABLE [dbo].[Valcar](
	[Recibo] [int] NULL,
	[Cliente] [char](6) NULL,
	[Cheque] [char](10) NOT NULL,
	[Banco] char(20) NULL,
	[Impo1] [float] NULL,
	[Impo2] [float] NULL,
	[Impo3] [float] NULL,
	[Impo4] [float] NULL,
	[Impo5] [float] NULL,
	[Titulo1] [char](10) NOT NULL,
	[Titulo2] [char](10) NOT NULL,
	[Titulo3] [char](10) NOT NULL,
	[Titulo4] [char](10) NOT NULL,
	[Titulo5] [char](10) NOT NULL,
	
) ON [PRIMARY]
GO





--
--   ---------------------------------------------------------------------------------------------------------
--


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_Valcar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_Valcar]
GO


CREATE PROCEDURE PR_alta_Valcar
	(@Recibo int,
	@Cliente char(6),
	@Cheque char(10),
	@Banco char(20),
	@Impo1 float,
	@Impo2 float,
	@Impo3 float,
	@Impo4 float,
	@Impo5 float,
	@Titulo1 char(10),
	@Titulo2 char(10),
	@Titulo3 char(10),
	@Titulo4 char(10),
	@Titulo5 char(10))
AS
	INSERT INTO dbo.Valcar
		(Recibo, Cliente, Cheque, Banco, Impo1, Impo2,Impo3, Impo4, Impo5, 
		Titulo1, Titulo2, Titulo3, Titulo4, Titulo5)
		VALUES
		(@Recibo, @Cliente, @Cheque, @Banco, @Impo1, @Impo2,@Impo3, @Impo4, @Impo5, 
		@Titulo1, @Titulo2, @Titulo3, @Titulo4, @Titulo5)
GO





--
--   ---------------------------------------------------------------------------------------------------------
--

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cheques_valcar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cheques_valcar]
GO

CREATE PROCEDURE [dbo].[PR_buscar_cheques_valcar]
	(@DesdeFecha char(8)
	, @HastaFecha char(8)
	, @DesdeCliente char(6)
	, @HastaCliente char(6))
AS
	select Recibos.FechaOrd as FechaOrd
		 , recibos.Recibo as Recibo
		 , recibos.TipoReg as TipoReg
		 , recibos.TipoRec as TipoRec
		 , recibos.Cuenta as Cuenta
		 , recibos.Cliente as Cliente
		 , recibos.Tipo1 as Tipo1
		 , recibos.letra1 as Letra1
		 , recibos.Punto1 as Punto1
		 , recibos.Numero1 as Numero1
		 , recibos.Fecha as Fecha
		 , recibos.Tipo2 as Tipo2
		 , recibos.Numero2 as Numero2
		 , recibos.Importe1 as Importe1
		 , recibos.Paridad as Paridad
		 , recibos.Importe2 as Importe2
		 , recibos.RetIva as RetIva
		 , recibos.RetOtra as RetOtra
		 , recibos.RetSuss as RetSuss
		 , recibos.RetGanancias as RetGanancias
		 , recibos.Renglon as Renglon
		 , recibos.banco2 AS Banco2
		 , recibos.Fecha2 AS Fecha2
		 , recibos.FechaOrd2 AS FechaOrd2
		 , recibos.Destino AS Destino
		 
	from surfactanSA.dbo.Recibos recibos
	WHERE recibos.Fechaord2 between @DesdeFecha and @HastaFecha
		and recibos.Cliente between @DesdeCliente and @HastaCliente
		and TipoReg = '2'
		and Tipo2 = '02'
		and Estado2 <> 'X'
	order by recibos.Clave


GO








--
--   ---------------------------------------------------------------------------------------------------------
--


IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_limpiar_Valcar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_limpiar_Valcar]
GO
CREATE PROCEDURE PR_limpiar_Valcar
AS
	DELETE 
	FROM dbo.Valcar
GO




--
--   ---------------------------------------------------------------------------------------------------------
--

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_buscar_cheques_valcar_cuit]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_buscar_cheques_valcar_cuit]
GO

CREATE PROCEDURE [dbo].[PR_buscar_cheques_valcar_cuit]
	(@DesdeFecha char(8)
	, @HastaFecha char(8)
	, @Cuit char(15))
AS
	select Recibos.FechaOrd as FechaOrd
		 , recibos.Recibo as Recibo
		 , recibos.TipoReg as TipoReg
		 , recibos.TipoRec as TipoRec
		 , recibos.Cuenta as Cuenta
		 , recibos.Cliente as Cliente
		 , recibos.Tipo1 as Tipo1
		 , recibos.letra1 as Letra1
		 , recibos.Punto1 as Punto1
		 , recibos.Numero1 as Numero1
		 , recibos.Fecha as Fecha
		 , recibos.Tipo2 as Tipo2
		 , recibos.Numero2 as Numero2
		 , recibos.Importe1 as Importe1
		 , recibos.Paridad as Paridad
		 , recibos.Importe2 as Importe2
		 , recibos.RetIva as RetIva
		 , recibos.RetOtra as RetOtra
		 , recibos.RetSuss as RetSuss
		 , recibos.RetGanancias as RetGanancias
		 , recibos.Renglon as Renglon
		 , recibos.banco2 AS Banco2
		 , recibos.Fecha2 AS Fecha2
		 , recibos.FechaOrd2 AS FechaOrd2
		 , recibos.Destino AS Destino
		 
	from surfactanSA.dbo.Recibos recibos
	WHERE recibos.Fechaord2 between @DesdeFecha and @HastaFecha
		and recibos.Cuit = @Cuit
		and TipoReg = '2'
		and Tipo2 = '02'
	order by recibos.Clave


GO






--
--   ---------------------------------------------------------------------------------------------------------
--

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_pagos_titulo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_pagos_titulo]
GO

CREATE PROCEDURE [dbo].[PR_modificar_pagos_titulo]
	(@titulo varchar(50)
	, @tituloi varchar(50)
	, @DesdeFecha char(8)
	, @HastaFecha char(8))
AS
BEGIN
	BEGIN TRAN
		UPDATE	Pagos
		SET titulo = @titulo
		  , tituloi = @tituloi
		WHERE Pagos.FechaOrd between @DesdeFecha and @HastaFecha
	COMMIT
END
GO
