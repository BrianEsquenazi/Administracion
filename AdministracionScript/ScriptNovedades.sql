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

	SELECT *
	FROM
	(
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
			
		UNION ALL
		
		SELECT 
			 rep.Estado2
			, rep.Importe2
			, rep.Numero2
			, rep.Fecha2
			, LTRIM(RTRIM(rep.banco2)) as banco2
			, rep.Clave
			, rep.FechaOrd2
			, 'RecibosProvisorios' AS Tabla
		FROM surfactanSA.dbo.RecibosProvi rep
		WHERE rep.TipoReg = 2
			AND rep.Estado2 = 'P'
			AND rep.Tipo2 = '02'
			AND ISNULL(rep.ReciboDefinitivo,0) = 0 
	)td
	ORDER BY td.FechaOrd2, td.Numero2

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

CREATE PROCEDURE PR_get_carga_intereses
AS
	SELECT ISNULL(ccp.FechaOriginal,'') FechaOriginal 
		, ISNULL(ccp.DesProveOriginal,'') DesProveOriginal
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
		and ISNULL(ccp.Saldo,0) <> 0
		and ISNULL(ccp.Interes,0) = 0
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