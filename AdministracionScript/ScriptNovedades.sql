USE [surfactanSA]
GO

/*
		ELIMINACION NOVEDADES
*/

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_cheque_en_cartera]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_cheque_en_cartera]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_recibos]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_recibos]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_clave_lectora]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_clave_lectora]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_deposito]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_deposito]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_recibos_marca]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_recibos_marca]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_carga_intereses]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_carga_intereses]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_deposito_por_clave]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_deposito_por_clave]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_carga_intereses]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_modificar_carga_intereses]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_iva_compra]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_iva_compra]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_imputacion]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_imputacion]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_iva_compras_adicional]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_iva_compras_adicional]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_cuenta_corriente]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_cuenta_corriente]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_compra_por_codigo]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_compra_por_codigo]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_imputaciones_por_nro_interno]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_imputaciones_por_nro_interno]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_factura_pagada]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_factura_pagada]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_cuentas_sin_saldar]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_cuentas_sin_saldar]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_cheques_terceros]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_cheques_terceros]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_cheques_propios]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_cheques_propios]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_pago_por_orden]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_pago_por_orden]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_pago_pago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_pago_pago]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_alta_pago_forma_de_pago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_alta_pago_forma_de_pago]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_siguiente_orden_pago]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_siguiente_orden_pago]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_get_siguiente_deposito]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[PR_get_siguiente_deposito]
GO


/*
		GENERACION NOVEDADES
*/

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO

CREATE PROCEDURE PR_get_cheque_en_cartera 
AS
BEGIN

	--SELECT *
	--FROM
	--(
		SELECT 
			 re.Estado2
			, re.Importe2
			, re.Numero2
			, re.Fecha2
			, LTRIM(RTRIM(re.banco2)) as banco2
			, re.Clave
			, re.FechaOrd2
			, 'Recibos' AS Tabla
		FROM surfactanSA.dbo.Recibos re
		WHERE re.TipoReg = 2
			AND re.Estado2 <> 'X'
			AND re.Tipo2 = '02'
		ORDER BY re.FechaOrd2, re.Numero2
			
	--	UNION ALL
		
	--	SELECT 
	--		 rep.Estado2
	--		, rep.Importe2
	--		, rep.Numero2
	--		, rep.Fecha2
	--		, LTRIM(RTRIM(rep.banco2)) as banco2
	--		, rep.Clave
	--		, rep.FechaOrd2
	--		, 'RecibosProvisorios' AS Tabla
	--	FROM surfactanSA.dbo.RecibosProvi rep
	--	WHERE rep.TipoReg = 2
	--		AND rep.Estado2 = 'P'
	--		AND rep.Tipo2 = '02'
	--		AND ISNULL(rep.ReciboDefinitivo,0) = 0 
	--)td
	--ORDER BY td.FechaOrd2, td.Numero2

END
GO

CREATE PROCEDURE PR_modificar_recibos 
	(@clave char(8)
	, @estado varchar(1)
	, @destino varchar(50)
	, @tabla varchar(20))
AS
BEGIN

	declare @sql varchar(max)
		
	SET @sql = 'UPDATE '
	
	
	SET @sql = @sql + CASE @tabla WHEN 'recibos' THEN 'Recibos'
									ELSE 'RecibosProvi'
						END
	
	SET @sql = @sql + ' SET estado2 = ''' + @estado + ''' , destino = ''' + @destino + ''' WHERE clave = ''' + @clave + ''';'
	
	exec (@sql) 

END
GO

CREATE PROCEDURE PR_get_clave_lectora 
	(@clave char(31)
	, @tabla varchar(20))
AS
BEGIN

	declare @sql varchar(max)
		
	SET @sql = 'SELECT clave, recibo FROM '
	
	
	SET @sql = @sql + CASE @tabla WHEN 'recibos' THEN 'Recibos'
									ELSE 'RecibosProvi'
						END
	
	SET @sql = @sql + ' WHERE claveCheque = ''' + @clave + ''';'
	
	exec (@sql) 

END
GO


CREATE procedure [dbo].[PR_alta_deposito]
	@Clave    VarChar(8),
	@Deposito VarChar(6),
	@Renglon  varchar(2),
	@Banco    smallint,
	@Fecha    VarChar(10),  
	@Importe  Float,                                             
	@Acredita VarChar(10),  
	@Tipo2    VarChar(2),
	@Numero2  VarChar(8),
	@Fecha2   VarChar(10),  
	@Importe2 real,                
	@Observaciones2 VarChar(20)
AS
BEGIN
	declare @fechaOrd varchar(8) = (select dbo.FN_get_fecha_ordenable (@Fecha))
	declare @AcreditaOrd VarChar(8) = (select dbo.FN_get_fecha_ordenable (@Acredita))
	INSERT INTO
		Depositos
			(
			Clave    ,
			Deposito ,
			Renglon ,
			Banco  ,
			Fecha   ,  
			FechaOrd ,
			Importe   ,                                           
			Acredita   ,
			AcreditaOrd ,
			Tipo2 ,
			Numero2, 
			Fecha2  ,  
			Importe2 ,               
			Observaciones2,      
			Empresa ,
			Impolista
			)
	VALUES
			(
			@Clave    ,
			@Deposito ,
			@Renglon ,
			@Banco  ,
			@Fecha   ,   
			@FechaOrd ,
			@Importe   ,                                            
			@Acredita   ,
			@AcreditaOrd ,
			@Tipo2 ,
			@Numero2,  
			@Fecha2  ,   
			@Importe2 ,                
			@Observaciones2,       
			1 ,
			0
			)
END
GO

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PR_modificar_recibos_marca]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE procedure [dbo].[PR_modificar_recibos_marca]
	@Recibo VarChar(6) ,
	@FechaDepo char(10)  ,
	@FechaDepoOrd char(8),
	@Marca char(1)
AS
UPDATE
	Recibos
SET
	Marca = @Marca  ,
	FechaDepo = @FechaDepo   ,
	FechaDepoOrd =  @FechaDepoOrd

WHERE
	Recibo = @Recibo' 
END
GO

CREATE PROCEDURE PR_get_carga_intereses (@tipo char(1))
AS
	SELECT ISNULL(ccp.FechaOriginal,'') FechaOriginal 
		, RTRIM(ISNULL(ccp.DesProveOriginal,'')) DesProveOriginal
		, ISNULL(ccp.FacturaOriginal,'') FacturaOriginal
		, ISNULL(ccp.Cuota,'') Cuota
		, ISNULL(ccp.fecha,'') fecha
		, ISNULL(ccp.Saldo,0) as Saldo
		, ISNULL(ccp.Interes,0) as Intereses
		, ISNULL(ccp.IvaInteres,0) as IvaIntereses
		, ISNULL(ccp.Referencia,'') as Referencia
		, ISNULL(ccp.Clave,'') Clave
		, ISNULL(ccp.NroInterno,'') NroInterno
	FROM CtaCtePrv ccp
	WHERE ccp.Proveedor = '10077777777'
		AND ((@tipo = 'C' 
			and ISNULL(ccp.Saldo,0) <> 0
			and ISNULL(ccp.Interes,0) = 0)
			OR
			(@tipo = 'M' 
			and ISNULL(ccp.Saldo,0) <> 0
			and ISNULL(ccp.Interes,0) <> 0))

	ORDER BY ccp.OrdFechaOriginal
GO

CREATE PROCEDURE [dbo].[PR_get_deposito_por_clave]
	@Clave Char(8)
 AS

SELECT *
FROM Depositos 
WHERE
	Clave = @Clave
GO

CREATE PROCEDURE [dbo].[PR_modificar_carga_intereses]
	(@clave varchar(26)
	, @saldo float
	, @intereses float
	, @ivaIntereses float
	, @referencia varchar(10)) 
AS
BEGIN
	declare @saldo_nuevo float = @saldo + @intereses + @ivaIntereses
	declare @nro_interno int = (select NroInterno from CtaCtePrv where Clave = @clave)

	BEGIN TRAN
		UPDATE	CtaCtePrv
		SET Saldo = @saldo_nuevo
			, Interes = @intereses 
			, IvaInteres = @ivaIntereses
			, Referencia = @referencia
		WHERE Clave = @clave
		
		UPDATE IvaComp
		SET Neto = @intereses
			, Iva21 = @ivaIntereses
		WHERE NroInterno = @nro_interno
			
	COMMIT
END
GO

CREATE PROCEDURE [dbo].[PR_alta_iva_compra]
	@NroInterno  int,
	@Proveedor char(11),
	@Tipo char(2),
	@Letra char(1),
	@Punto char(4),
	@Numero char(8),  
	@Fecha  char(10),    
	@Vencimiento char(10),
	@Vencimiento1 char(10),
	@Periodo    char(10),
	@Neto   float,                                               
	@Iva21  float,                                               
	@Iva5   float,                                               
	@Iva27  float,                                               
	@Ib     float,                                               
	@Exento float,                                               
	@Contado char(1),
	@Impre char(2),
	@Paridad float  ,
	@Pago  int ,    
	@cai varchar(14),
	@VtoCai  varchar(10),
	@Iva105 float, 
	@Despacho varchar(20), 
	@Remito varchar(30),
	@SoloIva int                               
AS
BEGIN

	DECLARE @Ordfecha varchar(8) = (SELECT dbo.FN_verificar_fecha_ordenable (@Fecha))	
	BEGIN TRAN
	
		IF (NOT EXISTS (SELECT 1 FROM IvaComp ic where ic.NroInterno = @NroInterno))
			INSERT INTO IvaComp
				(
				NroInterno, Proveedor, Tipo , Letra , Punto , Numero ,  
				Fecha, Vencimiento , Vencimiento1 , Periodo, Neto,                                          
				Iva21, Iva5, Iva27, Ib, Exento, Contado , Impre ,
				Ordfecha , Empresa , Netolist , ExentoList  , Paridad   ,
				Pago, Cai, VtoCai, Iva105, Despacho, Remito, SoloIva 		                                          
				)
			VALUES
				(	
				@NroInterno, @Proveedor, @Tipo, @Letra, @Punto, @Numero,  
				@Fecha, @Vencimiento, @Vencimiento1, @Periodo, @Neto,                                               
				@Iva21, @Iva5, @Iva27, @Ib, @Exento, @Contado, @Impre,
				@Ordfecha, 1, 0, 0, @Paridad    ,
				@Pago, @cai, @VtoCai, @Iva105, @Despacho, @Remito, @SoloIva 
				)
		ELSE
			UPDATE IvaComp
			SET
				Proveedor = @Proveedor,
				Tipo = @Tipo,
				Letra = @Letra,
				Punto = @Punto,
				Numero = @Numero,    
				Fecha = @Fecha,   
				Vencimiento = @Vencimiento,
				Vencimiento1 = @Vencimiento1,
				Periodo = @Periodo,
				Neto = @Neto,
				Iva21 = @Iva21,
				Iva5 = @Iva5,
				Iva27 = @Iva27,                                               
				Ib = @Ib,                                               
				Exento = @Exento,                                               
				Contado = @Contado,
				Impre = @Impre,
				Ordfecha = @Ordfecha,
				Paridad = @Paridad  ,
				Pago = @Pago,  
				Cai = @Cai,
				VtoCai = @VtoCai,
				Iva105 = @Iva105,
				Despacho = @Despacho,
				Remito = @Remito,
				SoloIva = @SoloIva	                                    
			WHERE
				NroInterno = @NroInterno
		
		EXEC PR_baja_imputaciones  @NroInterno = @NroInterno
 	COMMIT
END
GO

CREATE PROCEDURE [dbo].[PR_alta_imputacion]
	@Clave varchar(24),                    
	@TipoMovi varchar(1), 
	@Proveedor varchar(11),   
	@TipoComp varchar(2),
	@LetraComp varchar(1), 
	@PuntoComp varchar(4), 
	@NroComp varchar(8), 
	@Renglon varchar(2),
	@Fecha   varchar(10),   
	@Observaciones varchar(30),                  
	@Cuenta   varchar(10)  ,
	@Debito   float  ,                                           
	@Credito  float  ,                                                                                  
	@NroInterno int
AS
BEGIN

	DECLARE @FechaOrd varchar(8) = (SELECT dbo.FN_verificar_fecha_ordenable (@Fecha))
	
	INSERT INTO Imputac
		(
			Clave, TipoMovi, Proveedor, TipoComp, LetraComp, PuntoComp, 
			NroComp , Renglon, Fecha  , Observaciones, Cuenta, Debito,                                           
			Credito, FechaOrd , Titulo, Empresa , DebitoList,
			CreditoList, NroInterno  
		)	
	VALUES
		(
			@Clave, @TipoMovi, @Proveedor, @TipoComp, @LetraComp, @PuntoComp,
			@NroComp, @Renglon, @Fecha, @Observaciones, @Cuenta, @Debito,
			@Credito, @FechaOrd, 'Compras', 1, 0,
			0, @NroInterno
		)

END
GO

CREATE PROCEDURE PR_alta_iva_compras_adicional
	(@Clave varchar(10),
	@NroInterno int,
	@Renglon int,
	@Cuit varchar(15),
	@Razon varchar(50),
	@Tipo varchar(2),
	@Letra char(1),
	@Punto varchar(4),
	@Numero varchar(8),
	@Fecha varchar(10),
	@Neto float,
	@Iva21 float,
	@Iva27 float,
	@Iva105 float,
	@PerceIva float,
	@PerceIb float,
	@Exento float)
AS
BEGIN

	DECLARE @OrdFecha varchar(8) = (SELECT dbo.FN_verificar_fecha_ordenable (@Fecha))
	
	INSERT INTO IvaCompAdicional (
			Clave , NroInterno , Renglon , Cuit , Razon , Tipo ,
			Letra , Punto , Numero , Fecha , OrdFecha , Neto ,
			Iva21 , Iva27 , Iva105 , PerceIva , PerceIb , Exento)
		Values (
			@Clave, @NroInterno, @Renglon, @Cuit, @Razon , @Tipo,
			@Letra, @Punto , @Numero, @Fecha, @OrdFecha, @Neto,
			@Iva21, @Iva27, @Iva105, @PerceIva, @PerceIb, @Exento)

END
GO

CREATE PROCEDURE PR_alta_cuenta_corriente
	(@Contado varchar(1),                   
	@Proveedor varchar(11),   
	@Letra varchar(1), 
	@Tipo varchar(2),
	@Punto varchar(4),
	@Numero varchar(8),  
	@fecha   varchar(10),   
	@Vencimiento varchar(10), 
	@Vencimiento1 varchar(50),                                       
	@Total   float    ,  
	@Saldo float,                                                                                 
	@Impre varchar(2),                                          
	@NroInterno int,  
	@Paridad   float   ,
	@Pago  int)
AS
BEGIN

	DECLARE @OrdFecha varchar(8) = (SELECT dbo.FN_verificar_fecha_ordenable (@fecha))
	DECLARE @OrdVencimiento varchar(8) = (SELECT dbo.FN_verificar_fecha_ordenable (@Vencimiento))
	DECLARE @Clave varchar(26) = @Proveedor + @Letra + @Tipo + @Punto + @Numero
	
	BEGIN TRAN
	-- LO SIGUIENTE SOLO OCURRE DE SER CTA CTE / PYME NACION
		IF (@Contado = 2 OR @Contado = 3)
			IF(NOT EXISTS (SELECT 1 FROM CtaCtePrv cc WHERE	cc.Clave = @Clave))
				INSERT INTO CtaCtePrv
					(
						Clave , Proveedor, Letra, Tipo , Punto , Numero ,  fecha   ,   
						Estado , Vencimiento, Vencimiento1, Total, Saldo, OrdFecha ,
						OrdVencimiento, Impre , Empresa, SaldoList, NroInterno,  
						Lista , Acumulado , Paridad , Pago , Observaciones , Tarjeta  
					)	
				VALUES
					(
						@Clave , @Proveedor , @Letra , @Tipo, @Punto, @Numero, @fecha,   
						1, @Vencimiento, @Vencimiento1, @Total, @Saldo, @OrdFecha,
						@OrdVencimiento, @Impre, 1, 0, @NroInterno,  
						"", 0 , @Paridad , @Pago, "", @Contado
					)
			ELSE
				UPDATE CtaCteprv
				SET	Clave  =   @Clave ,
					Proveedor = @Proveedor,   
					Letra  = 	@Letra,
					Tipo = @Tipo ,
					Punto = @Punto ,
					Numero = @Numero ,
					fecha   = @fecha  ,
					Estado = 1,
					Vencimiento = @Vencimiento ,
					Vencimiento1 =   @Vencimiento1 ,
					Total   = @Total  ,
					Saldo  = @Saldo ,
					OrdFecha  = @OrdFecha  ,
					OrdVencimiento = @OrdVencimiento ,
					Impre = 	@Impre ,
					Empresa = 1, 
					SaldoList = 0,
					NroInterno = @NroInterno ,
					Lista  = 	"",
					Acumulado = 0,
					Paridad = @Paridad  ,
					Pago = @Pago , 
					Observaciones = "" , 
					Tarjeta = @Contado		
				WHERE
					Clave = @Clave

	COMMIT

END
GO

CREATE PROCEDURE PR_get_compra_por_codigo
	(@nro_interno int)
AS
	SELECT *
	FROM IvaComp ic
	WHERE ic.NroInterno = @nro_interno
GO

CREATE PROCEDURE PR_get_imputaciones_por_nro_interno
	(@nro_interno int)
AS
	SELECT *
	FROM Imputac im
	WHERE im.NroInterno = @nro_interno
GO

CREATE PROCEDURE PR_factura_pagada
	(@nro_interno int)
AS
	IF EXISTS (SELECT *
	FROM CtaCtePrv cta
	WHERE cta.NroInterno = @nro_interno
	AND cta.Saldo <> cta.Total)
	RETURN 1
	ELSE
	RETURN 0
GO

CREATE PROCEDURE PR_get_cuentas_sin_saldar
	(@Proveedor varchar(11))
AS
	SELECT *
	FROM CtaCtePrv cta
	WHERE
	cta.Proveedor = @Proveedor AND
	cta.Saldo <> 0
GO

CREATE PROCEDURE [dbo].[PR_get_cheques_terceros] (@cheque varchar(8)) AS

	SELECT 
		LTRIM(RTRIM(ISNULL(td.Numero2,''))) as Numero2
		, LTRIM(RTRIM(ISNULL(td.Banco,''))) as Banco
		, td.Importe2
		, LTRIM(RTRIM(ISNULL(td.Fecha,''))) as Fecha
		, LTRIM(RTRIM(ISNULL(td.Fecha2,''))) as Fecha2
		, LTRIM(RTRIM(ISNULL(td.Recibo,''))) as Recibo
		, LTRIM(RTRIM(ISNULL(td.Cliente,''))) as Cliente
	FROM
	(
		SELECT r.Numero2
			, r.Banco2 as Banco
			, r.Importe2
			, r.Fecha
			, r.Fecha2
			, r.Recibo
			, 'Rec: ' + cli.Razon as Cliente
			, r.FechaOrd2 as orden1
			, r.Numero2 as orden2
		FROM Recibos r
		JOIN Cliente cli on cli.Cliente = r.Cliente
		Where	Tiporeg= '2' 
			and  (Tipo2 = '2' or Tipo2 = '02')
			and ISNULL(Numero2,'') LIKE ('%' + @cheque + '')			
		
		UNION ALL
		
		SELECT d.Numero2
			, d.Observaciones2
			, d.Importe2
			, d.Fecha
			, d.Fecha2
			, d.Deposito
			, 'Dep: ' + LTRIM(RTRIM(ISNULL(b.Nombre,''))) 
			, d.Deposito 
			, 1 
		FROM Depositos d
		JOIN Banco b on b.Banco = d.Banco
		Where (Tipo2 = '3' Or Tipo2= '03')
			and ISNULL(Numero2,'') LIKE ('%' + @cheque + '')	
			
		UNION ALL
		
		SELECT p.Numero2
			, p.Observaciones2
			, p.Importe2
			, p.Fecha
			, p.Fecha2
			, p.Orden
			, 'O.P.: ' + LTRIM(RTRIM(ISNULL(pro.Nombre,p.Observaciones))) 
			, p.Orden
			, 1
		FROM Pagos p
		LEFT JOIN Proveedor pro on pro.Proveedor = p.Proveedor
		Where  Tiporeg = '2' 
			and (Tipo2 = '3' Or Tipo2= '03')
			and ISNULL(Numero2,'') LIKE ('%' + @cheque + '')	
	)td
	ORDER BY td.orden1, td.orden2
GO

CREATE PROCEDURE [dbo].[PR_get_cheques_propios] (@cheque varchar(8)) AS

	SELECT 	LTRIM(RTRIM(ISNULL(p.Numero2,''))) as Numero2
		, LTRIM(RTRIM(ISNULL(p.Observaciones2,''))) as Banco2
		, p.Importe2
		, LTRIM(RTRIM(ISNULL(p.Fecha,''))) as Fecha
		, LTRIM(RTRIM(ISNULL(p.Fecha2,''))) as Fecha2
		, LTRIM(RTRIM(ISNULL(p.Orden,''))) as Recibo
		, 'O.P.: ' + LTRIM(RTRIM(ISNULL(pro.Nombre,p.Observaciones))) as Proveedor 
	FROM Pagos p
	LEFT JOIN Proveedor pro on pro.Proveedor = p.Proveedor
	Where Tiporeg = '2' and Tipo2 = '02'
		and ISNULL(Numero2,'') LIKE ('%' + @cheque + '')
	ORDER BY p.Orden	

GO

CREATE PROCEDURE PR_get_pago_por_orden (@orden VARCHAR(6)) AS
	SELECT p.Orden
		, p.TipoOrd
		, p.Fecha
		, p.Proveedor
		, p.Observaciones
		, p.banco2 AS banco
		, p.Fecha2 AS fechaParidad
		, p.Paridad
		, p.RetGanancias
		, p.RetOtra AS RetencionIB
		, p.RetIbCiudad
		, p.RetIva
		, p.Importe
		, p.TipoReg
		, p.Tipo1
		, p.Letra1
		, p.Punto1
		, p.Numero1
		, p.Importe1
		, p.cuenta
		, p.Observaciones2
		, p.Tipo2
		, p.Numero2
		, p.FechaCheque
		, p.banco2
		, p.BancoCheque AS NombreCheque
		, p.Importe2
	FROM Pagos p
	WHERE p.Orden = @orden
GO

CREATE PROCEDURE PR_alta_pago_pago 
	(@Orden VARCHAR(6)
	, @Renglon varchar(2)
	, @TipoOrd varchar(1)
	, @Fecha varchar(10)
	, @Proveedor varchar(11)
 	, @Observaciones varchar(50)
	, @banco2 smallint
	, @Fecha2 varchar(10)
	, @Paridad float 
	, @RetGanancias real
	, @RetOtra real
	, @RetIbCiudad real
	, @RetIva real
	, @Importe real
	, @Tipo1 varchar(2)
	, @Letra1 varchar(1)
	, @Punto1 varchar(4)
	, @Numero1 varchar(8)
	, @Importe1 real
	, @cuenta varchar(10)
	, @Observaciones2 varchar(30)
	) 
AS
	DECLARE @Clave as varchar(8)
	SET @Clave = @Orden + @Renglon
	INSERT INTO Pagos
		(Clave, Orden, Renglon, TipoOrd, Fecha , Proveedor , Observaciones , banco2 , Fecha2 
	, Paridad ,RetGanancias , retOtra , RetIbCiudad , RetIva, Importe, Tipo1 , Letra1 
	, Punto1 , numero1 , Importe1 , cuenta, Observaciones2, TipoReg) 
	VALUES
			(@Clave, @Orden, @Renglon, @TipoOrd, @Fecha , @Proveedor , @Observaciones , @banco2 , @Fecha2 
	, @Paridad , @RetGanancias , @RetOtra , @RetIbCiudad , @RetIva, @Importe, @Tipo1 , @Letra1 
	, @Punto1 , @Numero1 , @Importe1 , @cuenta, @Observaciones2, 1 ) 
	
GO

CREATE PROCEDURE PR_alta_pago_forma_de_pago
	(@Orden VARCHAR(6)
	, @Renglon varchar(2)
	, @TipoOrd varchar(1)
	, @Fecha varchar(10)
	, @Proveedor varchar(11)
 	, @Observaciones varchar(50)
	, @banco2 smallint
	, @Fecha2 varchar(10)
	, @Paridad float 
	, @RetGanancias real
	, @RetOtra real
	, @RetIbCiudad real
	, @RetIva real
	, @Importe real
	, @Tipo2 varchar(2)
	, @Numero2 varchar(8)
	, @FechaCheque char(10)
	, @BancoCheque char(30)
	, @Importe2 float
	, @cuenta varchar(10)
	, @Observaciones2 varchar(30)
	) 
AS
	DECLARE @Clave as varchar(8)
	SET @Clave = @Orden + @Renglon
	INSERT INTO Pagos
		(Clave, Orden, Renglon, TipoOrd, Fecha , Proveedor , Observaciones , banco2 , Fecha2 
	, Paridad ,RetGanancias , retOtra , RetIbCiudad , RetIva, Importe, Tipo2 , Numero2 
	, FechaCheque , BancoCheque , Importe2 , cuenta, Observaciones2, TipoReg) 
	VALUES
			(@Clave, @Orden, @Renglon, @TipoOrd, @Fecha , @Proveedor , @Observaciones , @banco2 , @Fecha2 
	, @Paridad , @RetGanancias , @RetOtra , @RetIbCiudad , @RetIva, @Importe, @Tipo2 , @Numero2 
	, @FechaCheque , @BancoCheque , @Importe2 , @cuenta, @Observaciones2, 1 ) 

GO

CREATE PROCEDURE PR_get_siguiente_orden_pago
AS
	DECLARE @orden_siguiente VARCHAR(6)
	
	SET @orden_siguiente = (SELECT TOP 1 p.Orden
							FROM Pagos p
							ORDER BY p.Orden desc)
	
	RETURN (convert(int, @orden_siguiente) + 1)
GO

CREATE procedure PR_get_siguiente_deposito
AS
BEGIN
	
	declare @deposito_siguiente varchar(6)
	declare @clave_max varchar(8) = (SELECT MAX(clave) FROM Depositos)

	SELECT @deposito_siguiente = d.Deposito
	FROM Depositos d
	WHERE d.Clave = @clave_max
	
	return (CONVERT(int, @deposito_siguiente) + 1)
END
GO
