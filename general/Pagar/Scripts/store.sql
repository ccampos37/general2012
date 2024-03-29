if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TipoDoc]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[TipoDoc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_EMLB_PlanillaCobranza]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_EMLB_PlanillaCobranza]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_EMLB_PlanillaDocVarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_EMLB_PlanillaDocVarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_EMLB_SaldoxCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_EMLB_SaldoxCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_EMLB_SubPlanillaCobranza]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_EMLB_SubPlanillaCobranza]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_EMLB_SubPlanillaDocVarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_EMLB_SubPlanillaDocVarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_EMLB_SubSaldoxCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_EMLB_SubSaldoxCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_Ingresavarios_pro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_Ingresavarios_pro]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_SaldoxVendxCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_SaldoxVendxCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_abonadocumento_pro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_abonadocumento_pro]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_ingresacargo_pro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_ingresacargo_pro]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cp_ingresacargovalor_pro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cp_ingresacargovalor_pro]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE      PROC cp_EMLB_PlanillaCobranza 		/*EN USO*/
@base varchar(50),
@fecdesde varchar(10),
@fechasta varchar(10),
@codvendedor char(1)

AS

DECLARE @sensql nvarchar (4000)
SET @sensql = N'
SELECT 	e.abonocancli as Cod_Cliente,c.clienterazonsocial as Razon_Social,
	e.vendedorcodigo as Cod_Vendedor,e.documentoabono as Cod_Doc_Cargo,
	b.tdocumentodesccorta as Desc_Doc_Cargo,e.abononumdoc as Num_Doc_Cargo,
	a.cargoapefecemi as Fec_Emision_Cargo,
	e.abonocanforcan as Forma_Pago,d.monedasimbolo, 
	ISNULL(e.abonocanimpcan,0) as Importe_Abono,
	e.abonocantdqc as Cod_Doc_Abono,g.tdocumentodesccorta as Desc_Doc_Abono,
	e.abonocanndqc as Num_Doc_Abono,e.abonocanfecan as Fec_Cancela_Abono,
	e.abononumplanilla as Num_Planilla, e.abonotipoplanilla as Tipo_Planilla

FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b 
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN
	['+@base+'].dbo.cp_proveedor c
	ON
	a.clientecodigo = c.clientecodigo
	JOIN
	['+@base+'].dbo.gr_moneda d
	ON
	a.monedacodigo = d.monedacodigo
	JOIN
	['+@base+'].dbo.cp_abono e
	ON
	a.documentocargo = e.documentoabono
	AND a.cargonumdoc = e.abononumdoc
	JOIN
	['+@base+'].dbo.cp_tipodocumento g
	ON
	g.tdocumentocodigo = e.abonocantdqc
	JOIN
	['+@base+'].dbo.cp_tipoplanilla h
	ON
	h.tplanillacodigo = e.abonotipoplanilla
WHERE	
	e.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND e.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND h.tplanillacobranza = ''1''
	AND e.abonocanflreg IS NULL
ORDER BY 
	 e.abonotipoplanilla,e.abononumplanilla,e.abononumdoc,e.abonocanforcan '
exec (@sensql)
RETURN




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE      PROC cp_EMLB_PlanillaDocVarios 		/*EN USO*/
@base varchar(50),
@fecdesde varchar(10),
@fechasta varchar(10)

AS

DECLARE @sensql nvarchar (4000)

SET @sensql = N'
SELECT 	a.abononumplanilla as Num_Planilla,
	a.cargoapefecpla as Fec_Planilla, a.cargoapetipcam as Tipo_Cambio,
	a.clientecodigo as Cod_Cliente,c.clienterazonsocial as Razon_Social,
	a.documentocargo as Cod_Documento, b.tdocumentodesccorta as Desc_Documento,
	a.cargonumdoc as Num_Documento,a.cargoapefecemi as Fec_Emision,
	a.cargoapefecvct as Fec_Vencimiento,d.monedasimbolo,
	ISNULL(a.cargoapeimpape,0) as Importe_Apertura
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b 
	ON 
	a.documentocargo = b.tdocumentocodigo 
	JOIN 
	['+@base+'].dbo.cp_proveedor c 
	ON 
	a.clientecodigo = c.clientecodigo 
	JOIN 
	['+@base+'].dbo.gr_moneda d 
	ON 
	a.monedacodigo = d.monedacodigo 
	JOIN 
	['+@base+'].dbo.cp_tipoplanilla f 
	ON 
	f.tplanillacodigo = a.abonotipoplanilla 
WHERE	
	a.cargoapefecpla BETWEEN '''+@fecdesde+''' AND  '''+@fechasta+''' 
	AND f.tplanilladocvarios = ''1'' 
	AND a.cargoapeflgcan <> 1 
	AND a.cargoapeflgreg IS NULL
ORDER BY 
	a.abononumplanilla,d.monedasimbolo,a.cargoapefecpla,a.clientecodigo,a.cargonumdoc '
exec (@sensql)
RETURN











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE       PROC cp_EMLB_SaldoxCliente 		/*EN USO*/
@base varchar(50),
@ctacontable varchar(20),
@codcliente varchar(11),
@codmoneda varchar (2),
@fechasta varchar(10),
@letra bit

AS

DECLARE @sensql nvarchar (4000)
DECLARE @where nvarchar (4000)
SET @where = ''
IF @ctacontable <> '%'
BEGIN
	SET @where =  ' AND(LTRIM(RTRIM(f.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') OR LTRIM(RTRIM(f.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''')) '
	IF @codmoneda <> '%'
	    BEGIN
		IF @codmoneda = '01'
		   BEGIN
		     SET @where = ' AND LTRIM(RTRIM(f.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') '
		   END
		IF @codmoneda = '02'
	    	   BEGIN
		     SET @where = ' AND LTRIM(RTRIM(f.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''') '
	  	   END
	    END
END

IF convert(varchar(10),getdate(),103) = @fechasta
BEGIN
SET @sensql = N'
SELECT 	a.clientecodigo as Cod_Cliente,c.clienterazonsocial as Razon_Social,
	a.documentocargo as Cod_Documento,b.tdocumentodesccorta as Desc_Documento,
	a.cargonumdoc as Num_Documento,
	f.tdocumentocuentasoles as Cuenta_Soles, f.tdocumentocuentadolares as Cuenta_Dolar,
	a.cargoapefecemi as Fec_Emision,a.cargoapefecvct as Fec_Vencimiento,d.monedasimbolo,
	ISNULL( dbo.tipodoc(b.tdocumentotipo,a.cargoapeimpape) ,0 ) as Cargo,
	ISNULL( a.cargoapeimppag,0 ) as Abono
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b 
	ON 
	a.documentocargo = b.tdocumentocodigo 
	JOIN 
	['+@base+'].dbo.cp_proveedor c 
	ON 
	a.clientecodigo = c.clientecodigo 
	JOIN 
	['+@base+'].dbo.gr_moneda d 
	ON 
	a.monedacodigo = d.monedacodigo 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento f 
	ON 
	f.tdocumentocodigo = a.documentocargo 
WHERE	
	LTRIM(RTRIM(a.clientecodigo)) LIKE ('''+@codcliente+''') 
	AND a.cargoapefecemi <= '''+@fechasta+''' AND a.cargoapeflgcan <> 1 
	'+@where+' 
	AND a.monedacodigo LIKE ('''+@codmoneda+''') 
	AND a.cargoapeflgreg IS NULL
ORDER BY 
	a.clientecodigo,d.monedasimbolo,a.documentocargo,a.cargonumdoc '
exec (@sensql)
RETURN
END
ELSE
BEGIN
SET @sensql = N'
SELECT 	a.clientecodigo as Cod_Cliente,c.clienterazonsocial as Razon_Social,
	a.documentocargo as Cod_Documento,b.tdocumentodesccorta as Desc_Documento,
	a.cargonumdoc as Num_Documento,
	f.tdocumentocuentasoles as Cuenta_Soles,f.tdocumentocuentadolares as Cuenta_Dolar,
	a.cargoapefecemi as Fec_Emision,a.cargoapefecvct as Fec_Vencimiento,d.monedasimbolo,  	
	ISNULL( dbo.tipodoc(b.tdocumentotipo,a.cargoapeimpape) ,0 ) as Cargo,
	SUM( ISNULL( e.abonocanimpcan,0 ) ) as Abono
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN
	['+@base+'].dbo.cp_proveedor c
	ON
	a.clientecodigo = c.clientecodigo
	JOIN
	['+@base+'].dbo.gr_moneda d
	ON
	a.monedacodigo = d.monedacodigo
	JOIN
	['+@base+'].dbo.cp_abono e
	ON
	a.documentocargo = e.documentoabono
	AND a.cargonumdoc = e.abononumdoc
	JOIN
	['+@base+'].dbo.cp_tipodocumento f
	ON
	f.tdocumentocodigo = a.documentocargo
WHERE	
	LTRIM(RTRIM(a.clientecodigo)) LIKE ('''+@codcliente+''')
	AND a.monedacodigo LIKE ('''+@codmoneda+''')
	AND a.cargoapefecemi <= '''+@fechasta+'''
	AND a.cargoapeflgcan <> 1 
	'+@where+'
	AND CONVERT(varchar(10),e.abonocanfecpla,103) <= '''+@fechasta+'''
	AND a.cargoapeflgreg IS NULL
	  	
GROUP BY
	a.clientecodigo,c.clienterazonsocial,
	a.documentocargo,b.tdocumentodesccorta,
	a.cargonumdoc,
	f.tdocumentocuentasoles,f.tdocumentocuentadolares,
	a.cargoapefecemi,a.cargoapefecvct,d.monedasimbolo,  	
	b.tdocumentotipo,a.cargoapeimpape
ORDER BY 
	a.clientecodigo,d.monedasimbolo,a.documentocargo,a.cargonumdoc '
exec (@sensql)
RETURN
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  PROC cp_EMLB_SubPlanillaCobranza  		/*EN USO*/
@base varchar(50),
@fecdesde varchar(10),
@fechasta varchar(10),
@codvendedor char(1)

AS

DECLARE @sensql nvarchar (4000)
SET @sensql = N'
SELECT 	e.abonocantdqc as Cod_Doc_Abono,g.tdocumentodescripcion as Desc_Doc_Abono,
	e.documentoabono as Cod_Doc_Cargo,i.tdocumentodescripcion as Desc_Doc_Cargo,
	IMPORTES_DOLARES = 
	isnull ( (
	SELECT SUM (isnull(z.abonocanimpcan,0)) 
	FROM ['+@base+'].dbo.cp_abono z
	JOIN
	['+@base+'].dbo.cp_tipoplanilla y
	ON
	y.tplanillacodigo = z.abonotipoplanilla
	WHERE z.abonocanmoneda = 02 
	AND z.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND z.abonocanflreg IS NULL
	AND z.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND y.tplanillacobranza = ''1'' 
	AND e.abonocantdqc = z.abonocantdqc AND e.documentoabono = z.documentoabono
	) , 0 ) , 
	IMPORTES_SOLES = 
	isnull (  (
	SELECT SUM (isnull(x.abonocanimpcan,0)) 
	FROM ['+@base+'].dbo.cp_abono x
	JOIN
	['+@base+'].dbo.cp_tipoplanilla w
	ON
	w.tplanillacodigo = x.abonotipoplanilla
	WHERE x.abonocanmoneda = 01 
	AND x.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND x.abonocanflreg IS NULL
	AND x.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND w.tplanillacobranza = ''1'' 
	AND e.abonocantdqc = x.abonocantdqc AND e.documentoabono = x.documentoabono
	) , 0 )
FROM 	
	['+@base+'].dbo.cp_abono e
	JOIN
	['+@base+'].dbo.cp_tipodocumento g
	ON
	g.tdocumentocodigo = e.abonocantdqc
	JOIN
	['+@base+'].dbo.cp_tipoplanilla h
	ON
	h.tplanillacodigo = e.abonotipoplanilla
	JOIN
	['+@base+'].dbo.cp_tipodocumento i
	ON
	i.tdocumentocodigo = e.documentoabono
WHERE	
	e.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND e.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND h.tplanillacobranza = ''1''
	AND e.abonocanflreg IS NULL
GROUP BY 
	e.abonocantdqc,g.tdocumentodescripcion,
	e.documentoabono,i.tdocumentodescripcion
ORDER BY 
	g.tdocumentodescripcion,e.abonocantdqc '	
exec (@sensql)
RETURN




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE    PROC cp_EMLB_SubPlanillaDocVarios 		/*EN USO*/
@base varchar(50),
@fecdesde varchar(10),
@fechasta varchar(10)

AS

DECLARE @sensql nvarchar (4000)

SET @sensql = N'
SELECT 	a.documentocargo as Cod_Documento, b.tdocumentodescripcion as Desc_Documento,
	TOTAL_SOLES = CASE 
	WHEN a.monedacodigo = ''01'' THEN SUM(isnull(a.cargoapeimpape,0)) 
	ELSE 0
	end,
	TOTAL_DOLARES = CASE 
	WHEN a.monedacodigo = ''02'' THEN SUM(isnull(a.cargoapeimpape,0)) 
	ELSE 0
	end
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b 
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN 
	['+@base+'].dbo.cp_tipoplanilla f 
	ON 
	f.tplanillacodigo = a.abonotipoplanilla
WHERE	
	a.cargoapefecpla BETWEEN '''+@fecdesde+''' AND  '''+@fechasta+''' 
	AND f.tplanilladocvarios = ''1'' 
	AND a.cargoapeflgcan <> 1 
	AND a.cargoapeflgreg IS NULL
GROUP BY
	a.documentocargo,b.tdocumentodescripcion,a.monedacodigo 
ORDER BY 
	b.tdocumentodescripcion '
exec (@sensql)
RETURN







GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE proc cp_EMLB_SubSaldoxCliente 		/*EN USO*/
@base varchar(50),
@ctacontable varchar(20),
@codcliente varchar(11),
@codmoneda varchar (2),
@fechasta varchar(10),
@letra bit

AS

DECLARE @sensql nvarchar (4000)
DECLARE @where nvarchar (4000)
SET @where = ''
IF @ctacontable <> '%'
BEGIN
	SET @where =  ' AND(LTRIM(RTRIM(f.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') OR LTRIM(RTRIM(f.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''')) '
	IF @codmoneda <> '%'
	    BEGIN
		IF @codmoneda = '01'
		   BEGIN
		     SET @where = ' AND LTRIM(RTRIM(f.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') '
		   END
		IF @codmoneda = '02'
	    	   BEGIN
		     SET @where = ' AND LTRIM(RTRIM(f.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''') '
	  	   END
	    END
END

IF convert(varchar(10),getdate(),103) = @fechasta
BEGIN
SET @sensql = N'
SELECT 	a.documentocargo as Cod_Documento,b.tdocumentodescripcion as Desc_Documento,
	TOTAL_SOLES = CASE 
	WHEN a.monedacodigo = 01 THEN SUM(isnull(a.cargoapeimpape,0)) - SUM(isnull(a.cargoapeimppag,0) )
	ELSE 0
	end,
	TOTAL_DOLARES = CASE 
	WHEN a.monedacodigo = 02 THEN SUM(isnull(a.cargoapeimpape,0)) - SUM(isnull(a.cargoapeimppag,0))
	ELSE 0
	end
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b 
	ON 
	a.documentocargo = b.tdocumentocodigo 	
WHERE	
	LTRIM(RTRIM(a.clientecodigo)) LIKE ('''+@codcliente+''') 
	AND a.cargoapefecemi <= '''+@fechasta+''' AND a.cargoapeflgcan <> 1 
	'+@where+' 
	AND a.monedacodigo LIKE ('''+@codmoneda+''') 
	AND a.cargoapeflgreg IS NULL
GROUP BY
	a.documentocargo,b.tdocumentodescripcion,a.monedacodigo
ORDER BY 
	b.tdocumentodescripcion,a.documentocargo,a.monedacodigo '

--	JOIN 
--	['+@base+'].dbo.cp_tipodocumento f 
--	ON 
--	f.tdocumentocodigo = a.documentocargo 

exec (@sensql)
RETURN
END
ELSE
BEGIN
SET @sensql = N'
SELECT 	a.documentocargo as Cod_Documento,b.tdocumentodescripcion as Desc_Documento,
	TOTAL_SOLES = CASE 
	WHEN a.monedacodigo = 01 THEN SUM(isnull(a.cargoapeimpape,0)) - SUM(isnull(e.abonocanimpcan,0))
	ELSE 0
	end,
	TOTAL_DOLARES = CASE 
	WHEN a.monedacodigo = 02 THEN SUM(isnull(a.cargoapeimpape,0)) - SUM(isnull(e.abonocanimpcan,0))
	ELSE 0
	end
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN
	['+@base+'].dbo.cp_abono e
	ON
	a.documentocargo = e.documentoabono
	AND a.cargonumdoc = e.abononumdoc
WHERE	
	LTRIM(RTRIM(a.clientecodigo)) LIKE ('''+@codcliente+''')
	AND a.monedacodigo LIKE ('''+@codmoneda+''')
	AND a.cargoapefecemi <= '''+@fechasta+'''
	AND a.cargoapeflgcan <> 1 
	'+@where+'
	AND CONVERT(varchar(10),e.abonocanfecpla,103) <= '''+@fechasta+'''	  	
	AND a.cargoapeflgreg IS NULL
GROUP BY
	a.documentocargo,b.tdocumentodescripcion,a.monedacodigo
ORDER BY 
	b.tdocumentodescripcion,a.documentocargo,a.monedacodigo'

--	JOIN
--	['+@base+'].dbo.cp_tipodocumento f
--	ON
--	f.tdocumentocodigo = a.documentocargo

exec (@sensql)
RETURN
END
















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROC cp_ingresavarios_pro
@base varchar(50),
@tipo char(1),
@tabla varchar(50),
@tipodocu char(2),
@numero  char(11),
@cliente  char(11),
@vendedor char(3),
@zona 	char(3),
@apefecemi datetime,
@moneda	char(2),
@apeimppag float,
@usuario char(8),
@tipocambio float,
@fechaact datetime,
@flagcancel bit,
@tipoplanilla char(2),
@planilla char(6),
@vencimiento datetime,
@fechaplani datetime,
@banco char(2),
@cargoabono char(1)
As
Declare @cadena as nvarchar(4000)
Declare @parame as nvarchar(4000)

if @tipo='1' 
   Begin
	SET @cadena =N'Insert Into ['+@base +'].dbo.'+@tabla +
           		      '(documentocargo,
				cargonumdoc,
				clientecodigo,
				vendedorcodigo,
				zonacodigo,
				cargoapefecemi,
				monedacodigo,
				usuariocodigo,
				cargoapetipcam,
				fechaact,
				cargoapeflgcan,
	 		        cargoapeimpape,
				abonotipoplanilla,
				abononumplanilla,
				cargoapefecvct,
				cargoapefecpla,
				bancocodigo,
				cargoapecarabo)
                          VALUES (
				@tipodocu, 
				@numero,
				@cliente,
				@vendedor,
				@zona,
				@apefecemi,
				@moneda,
				@usuario,
				@tipocambio,
				@fechaact,
				@flagcancel,
				@apeimppag,
				@tipoplanilla,
				@planilla,
				@vencimiento,
				@fechaplani,
				@banco,
				@cargoabono)'

	SET @Parame = N'@tipodocu char(2),
			@numero  char(11),
			@cliente  char(11),
			@vendedor char(3),
			@zona 	char(3),
			@apefecemi datetime,
			@moneda	char(2),
			@apeimppag float,
			@usuario	char(8),
			@tipocambio	float,
			@fechaact	datetime,
			@flagcancel   bit,
			@tipoplanilla char(2),
			@planilla char(6),
			@vencimiento datetime,
			@fechaplani datetime,
			@banco char(2),
			@cargoabono char(1)'

	EXEC sp_executesql @cadena,@parame,
				@tipodocu, 
				@numero,
				@cliente,
				@vendedor,
				@zona,
				@apefecemi,
				@moneda,
				@apeimppag,
				@usuario,
				@tipocambio,
				@fechaact,
				@flagcancel,
				@tipoplanilla,
				@planilla,
				@vencimiento,
				@fechaplani,
				@banco,
				@cargoabono

	end 
if @tipo='2' 
  Begin
	SET @cadena =N'UPDATE ['+@base +'].dbo.'+@tabla +
           		 ' SET 
			   clientecodigo=@cliente,
			   vendedorcodigo=@vendedor,
			   zonacodigo=@zona,
			   cargoapefecemi=@apefecemi,
			   monedacodigo=@moneda,
			   cargoapeimpape=@apeimppag,
			   usuariocodigo=@usuario,
			   cargoapetipcam=@tipocambio,
			   fechaact=@fechaact,
		           cargoapeflgcan=@flagcancel,
			   abonotipoplanilla=@tipoplanilla,
			   abononumplanilla=@planilla,
			   cargoapefecvct=@vencimiento,
			   cargoapefecpla=@fechaplani,
			   bancocodigo=@banco,
			   cargoapecarabo=@cargoabono	
                         Where  documentocargo=@tipodocu and cargonumdoc=@numero'

				
	SET @Parame = N'@tipodocu char(2),
			@numero  char(11),
			@cliente  char(11),
			@vendedor char(3),
			@zona 	char(3),
			@apefecemi datetime,
			@moneda	char(2),
			@apeimppag float,
			@usuario char(8),
			@tipocambio float,
			@fechaact  datetime,
                        @flagcancel  bit,
			@tipoplanilla char(2),
			@planilla char(6),
			@vencimiento datetime,
			@fechaplani datetime,
			@banco char(2),
			@cargoabono char(1)'

	EXEC sp_executesql @cadena,@parame,@tipodocu, 
						@numero,
						@cliente,
						@vendedor,
						@zona,
						@apefecemi,
						@moneda,
						@apeimppag,
						@usuario,
						@tipocambio,
						@fechaact,
						@flagcancel,
						@tipoplanilla,
						@planilla,
						@vencimiento,
						@fechaplani,
						@banco,
						@cargoabono


 end     



















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE   cp_SaldoxVendxCliente 		/*EN USO*/
@base varchar(50),
@ctacontable varchar(20),
@vendedor varchar(04),
@CodCliente varchar(11),
@codmoneda varchar (2),
@fechasta varchar(10),
@letra bit

AS

DECLARE @sensql nvarchar (4000)
DECLARE @where nvarchar (4000)
SET @where = ''
IF @ctacontable <> '%'
BEGIN
	SET @where =  ' AND(LTRIM(RTRIM(f.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') OR LTRIM(RTRIM(f.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''')) '
	IF @codmoneda <> '%'
	    BEGIN
		IF @codmoneda = '01'
		   BEGIN
		     SET @where = ' AND LTRIM(RTRIM(f.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') '
		   END
		IF @codmoneda = '02'
	    	   BEGIN
		     SET @where = ' AND LTRIM(RTRIM(f.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''') '
	  	   END
	    END
END

IF convert(varchar(10),getdate(),103) = @fechasta
BEGIN
SET @sensql = N'
SELECT 	a.vendedorCodigo as Cpd_Vendedor,a.clientecodigo as Cod_Cliente,c.clienterazonsocial as Razon_Social,
	a.documentocargo as Cod_Documento,b.tdocumentodesccorta as Desc_Documento,
	a.cargonumdoc as Num_Documento,
	f.tdocumentocuentasoles as Cuenta_Soles, f.tdocumentocuentadolares as Cuenta_Dolar,
	a.cargoapefecemi as Fec_Emision,a.cargoapefecvct as Fec_Vencimiento,d.monedasimbolo,
	ISNULL( dbo.tipodoc(b.tdocumentotipo,a.cargoapeimpape) ,0 ) as Cargo,
	ISNULL( a.cargoapeimppag,0 ) as Abono
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b 
	ON 
	a.documentocargo = b.tdocumentocodigo 
	JOIN 
	['+@base+'].dbo.cp_proveedor c 
	ON 
	a.clientecodigo = c.clientecodigo 
	JOIN 
	['+@base+'].dbo.gr_moneda d 
	ON 
	a.monedacodigo = d.monedacodigo 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento f 
	ON 
	f.tdocumentocodigo = a.documentocargo 
             JOIN
	['+@base+'].dbo.cp_oficina g 
	ON 
	a.documentocargo=g.vendedorcodigo 
WHERE	
	LTRIM(RTRIM(a.clientecodigo)) LIKE ('''+@codcliente+''') 
	AND a.cargoapefecemi <= '''+@fechasta+''' AND a.cargoapeflgcan <> 1 
	'+@where+' 
	AND a.monedacodigo LIKE ('''+@codmoneda+''') 
	AND a.cargoapeflgreg IS NULL
ORDER BY 
	a.vendedorCodigo,a.clientecodigo,d.monedasimbolo,a.documentocargo,a.cargonumdoc '
exec (@sensql)
RETURN
END
ELSE
BEGIN
SET @sensql = N'
SELECT      a.clientecodigo as Cod_Cliente,c.clienterazonsocial as Razon_Social,
	a.documentocargo as Cod_Documento,b.tdocumentodesccorta as Desc_Documento,
	a.cargonumdoc as Num_Documento,
	f.tdocumentocuentasoles as Cuenta_Soles,f.tdocumentocuentadolares as Cuenta_Dolar,
	a.cargoapefecemi as Fec_Emision,a.cargoapefecvct as Fec_Vencimiento,d.monedasimbolo,  	
	ISNULL( dbo.tipodoc(b.tdocumentotipo,a.cargoapeimpape) ,0 ) as Cargo,
	SUM( ISNULL( e.abonocanimpcan,0 ) ) as Abono
FROM 	
	['+@base+'].dbo.cp_cargo a 
	JOIN 
	['+@base+'].dbo.cp_tipodocumento b
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN
	['+@base+'].dbo.cp_proveedor c
	ON
	a.clientecodigo = c.clientecodigo
	JOIN
	['+@base+'].dbo.gr_moneda d
	ON
	a.monedacodigo = d.monedacodigo
	JOIN
	['+@base+'].dbo.cp_abono e
	ON
	a.documentocargo = e.documentoabono
	AND a.cargonumdoc = e.abononumdoc
	JOIN
	['+@base+'].dbo.cp_tipodocumento f
	ON
	f.tdocumentocodigo = a.documentocargo
WHERE	
	LTRIM(RTRIM(a.clientecodigo)) LIKE ('''+@codcliente+''')
	AND a.monedacodigo LIKE ('''+@codmoneda+''')
	AND a.cargoapefecemi <= '''+@fechasta+'''
	AND a.cargoapeflgcan <> 1 
	'+@where+'
	AND CONVERT(varchar(10),e.abonocanfecpla,103) <= '''+@fechasta+'''
	AND a.cargoapeflgreg IS NULL
	  	
GROUP BY
	a.clientecodigo,c.clienterazonsocial,
	a.documentocargo,b.tdocumentodesccorta,
	a.cargonumdoc,
	f.tdocumentocuentasoles,f.tdocumentocuentadolares,
	a.cargoapefecemi,a.cargoapefecvct,d.monedasimbolo,  	
	b.tdocumentotipo,a.cargoapeimpape
ORDER BY 
	a.clientecodigo,d.monedasimbolo,a.documentocargo,a.cargonumdoc '
exec (@sensql)
RETURN
END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE procedure cp_abonadocumento_pro
@base varchar(50),
@tipo char(1),
@documentoabono char(2),
@abononumdoc char(11),
@abonocannumpag char(2),
@zonacodigo char(3),
@tipoplanilla char(2),
@vendedor char(3),
@numplanilla char(6),
@fechapla datetime,
@fechapro datetime,
@moneda char(2),
@abonocancarabo char(1),
@cuenta varchar(20),
@banco char(2),
@tipocam float,
@abonoflpres char(1),
@abonocanimpcan float,
@abonocanimpsol float,
@usuario char(8),
@fechaact datetime,
@forma char(1),
@monedacan char(2),
@abonocantd char(2),
@abonocannro char(11),
@fechacan datetime,
@cliente char(11)
As
Declare @ncadena as nvarchar(4000)
Declare @parame  as nvarchar(4000)

if @tipo=1
  Begin 
   
   set @ncadena=N'Insert Into ['+@base+'].dbo.cp_abono
	               (documentoabono,
			abononumdoc,
			abonocannumpag,
			zonacodigo,
			abonotipoplanilla,
			vendedorcodigo,
			abononumplanilla,
			abonocanfecpla,
			abonocanfecpro,
			abonocanmoneda,
			abonocancarabo,
			abonocancuenta,
			abonocanbco,
			abonocantipcam,	
			abonocanimpcan,
			abonocanimcan,
			abonocanimpsol,
			usuariocodigo,
			fechaact,
			abonocanforcan,
			abonocanmoncan,
			abonocantdqc,
			abonocanndqc,
			abonocanfecan,
			abonocancli,
			abonocanflpres)
		  Values (
			@documentoabono,
			@abononumdoc,
			@abonocannumpag,
			@zonacodigo,
			@tipoplanilla,
			@vendedor,
			@numplanilla,
			@fechapla,
			@fechapro,
			@moneda,
			@abonocancarabo,
			@cuenta,
			@banco,
			@tipocam,
			@abonocanimpcan,
			@abonocanimpcan,
			@abonocanimpsol,
			@usuario,
			@fechaact,
			@forma,
			@monedacan,
			@abonocantd,
			@abonocannro,
			@fechacan,
			@cliente,
			@abonoflpres)'

	set @parame=N'@documentoabono char(2),
			@abononumdoc char(11),
			@abonocannumpag char(2),
			@zonacodigo char(3),
			@tipoplanilla char(2),
			@vendedor char(3),
			@numplanilla char(6),
			@fechapla datetime,
			@fechapro datetime,
			@moneda char(2),
			@abonocancarabo char(1),
			@cuenta varchar(20),
			@banco char(2),
			@tipocam float,		
			@abonocanimpcan float,
			@abonocanimpsol float,
			@usuario char(8),
			@fechaact datetime,
			@forma char(1),
			@monedacan char(2),
			@abonocantd char(2),
			@abonocannro char(11),
			@fechacan datetime,
			@cliente char(11),
			@abonoflpres char(1)'

	Exec sp_executesql @ncadena,@parame,@documentoabono,
						@abononumdoc,
						@abonocannumpag,
						@zonacodigo,
						@tipoplanilla,
						@vendedor,
						@numplanilla,
						@fechapla,
						@fechapro,
						@moneda,
						@abonocancarabo,
						@cuenta,
						@banco,
						@tipocam,						
						@abonocanimpcan,
						@abonocanimpsol,
						@usuario,
						@fechaact,
						@forma,
						@monedacan,
						@abonocantd,
						@abonocannro,
						@fechacan,
						@cliente,
						@abonoflpres

   end	



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROC cp_ingresacargo_pro
@base varchar(50),
@tipo char(1),
@tabla varchar(50),
@tipodocu char(2),
@numero  char(11),
@cliente  char(11),
@vendedor char(3),
@zona 	char(3),
@apefecemi datetime,
@moneda	char(2),
@apeimppag float,
@usuario char(8),
@tipocambio float,
@fechaact datetime,
@flagcancel bit,
@cargoabono char(1),
@concepto char(2)
As
Declare @cadena as nvarchar(4000)
Declare @parame as nvarchar(4000)

if @tipo='1' 
   Begin
	SET @cadena =N'Insert Into ['+@base +'].dbo.'+@tabla +
           		      '(documentocargo,
				cargonumdoc,
				clientecodigo,
				vendedorcodigo,
				zonacodigo,
				cargoapefecemi,
				monedacodigo,
				usuariocodigo,
				cargoapetipcam,
				fechaact,
				cargoapeflgcan,
	 		        cargoapeimpape,
				cargoapecarabo,
				conceptocodigo)
                          VALUES (
				@tipodocu, 
				@numero,
				@cliente,
				@vendedor,
				@zona,
				@apefecemi,
				@moneda,
				@usuario,
				@tipocambio,
				@fechaact,
				@flagcancel,
				@apeimppag,
				@cargoabono,
				@concepto)'

	SET @Parame = N'@tipodocu char(2),
			@numero  char(11),
			@cliente  char(11),
			@vendedor char(3),
			@zona 	char(3),
			@apefecemi datetime,
			@moneda	char(2),
			@apeimppag float,
			@usuario	char(8),
			@tipocambio	float,
			@fechaact	datetime,
			@flagcancel   bit,
			@cargoabono char(1),
			@concepto char(2)'

	EXEC sp_executesql @cadena,@parame,
				@tipodocu, 
				@numero,
				@cliente,
				@vendedor,
				@zona,
				@apefecemi,
				@moneda,
				@apeimppag,
				@usuario,
				@tipocambio,
				@fechaact,
				@flagcancel,
				@cargoabono,
				@concepto
	end 
if @tipo='2' 
  Begin
	SET @cadena =N'UPDATE ['+@base +'].dbo.'+@tabla +
           		 ' SET 
			   clientecodigo=@cliente,
			   vendedorcodigo=@vendedor,
			   zonacodigo=@zona,
			   cargoapefecemi=@apefecemi,
			   monedacodigo=@moneda,
			   cargoapeimpape=@apeimppag,
			   usuariocodigo=@usuario,
			   cargoapetipcam=@tipocambio,
			   fechaact=@fechaact,
		           cargoapeflgcan=@flagcancel,
			   cargoapecarabo=@cargoabono,
			   conceptocodigo=@concepto
                         Where  documentocargo=@tipodocu and cargonumdoc=@numero'
				
	SET @Parame = N'@tipodocu char(2),
			@numero  char(11),
			@cliente  char(11),
			@vendedor char(3),
			@zona 	char(3),
			@apefecemi datetime,
			@moneda	char(2),
			@apeimppag float,
			@usuario char(8),
			@tipocambio float,
			@fechaact  datetime,
                        @flagcancel  bit,
			@cargoabono char(1),
			@concepto char(2)'

	EXEC sp_executesql @cadena,@parame,@tipodocu, 
						@numero,
						@cliente,
						@vendedor,
						@zona,
						@apefecemi,
						@moneda,
						@apeimppag,
						@usuario,
						@tipocambio,
						@fechaact,
						@flagcancel,
						@cargoabono,
						@concepto

 end     
















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROC cp_ingresacargovalor_pro
@base varchar(50),
@tipo char(1),
@tabla varchar(50),
@tipodocu char(2),
@numero  char(11),
@cliente  char(11),
@vendedor char(3),
@zona 	char(3),
@apefecemi datetime,
@moneda	char(2),
@apeimppag float,
@usuario char(8),
@tipocambio float,
@fechaact datetime,
@flagcancel bit,
@cargoabono char(1),
@referencia varchar(80),
@concepto char(2),
@venci datetime
As
Declare @cadena as nvarchar(4000)
Declare @parame as nvarchar(4000)

if @tipo='1' 
   Begin
	SET @cadena =N'Insert Into ['+@base +'].dbo.'+@tabla +
           		      '(documentocargo,
				cargonumdoc,
				clientecodigo,
				vendedorcodigo,
				zonacodigo,
				cargoapefecemi,
				monedacodigo,
				usuariocodigo,
				cargoapetipcam,
				fechaact,
				cargoapeflgcan,
	 		        cargoapeimpape,
				cargoapecarabo,
				cargoaperefere,
				conceptocodigo,
				cargoapefecvct)
                          VALUES (
				@tipodocu, 
				@numero,
				@cliente,
				@vendedor,
				@zona,
				@apefecemi,
				@moneda,
				@usuario,
				@tipocambio,
				@fechaact,
				@flagcancel,
				@apeimppag,
				@cargoabono,
				@referencia,
				@concepto,
				@venci)'

	SET @Parame = N'@tipodocu char(2),
			@numero  char(11),
			@cliente  char(11),
			@vendedor char(3),
			@zona 	char(3),
			@apefecemi datetime,
			@moneda	char(2),
			@apeimppag float,
			@usuario	char(8),
			@tipocambio	float,
			@fechaact	datetime,
			@flagcancel   bit,
			@cargoabono char(1),
			@referencia varchar(200),
			@concepto char(2),
			@venci datetime'

	EXEC sp_executesql @cadena,@parame,
				@tipodocu, 
				@numero,
				@cliente,
				@vendedor,
				@zona,
				@apefecemi,
				@moneda,
				@apeimppag,
				@usuario,
				@tipocambio,
				@fechaact,
				@flagcancel,
				@cargoabono,
				@referencia,
				@concepto,
				@venci
	end 
if @tipo='2' 
  Begin
	SET @cadena =N'UPDATE ['+@base +'].dbo.'+@tabla +
           		 ' SET 
			   clientecodigo=@cliente,
			   vendedorcodigo=@vendedor,
			   zonacodigo=@zona,
			   cargoapefecemi=@apefecemi,
			   monedacodigo=@moneda,
			   cargoapeimpape=@apeimppag,
			   usuariocodigo=@usuario,
			   cargoapetipcam=@tipocambio,
			   fechaact=@fechaact,
		           cargoapeflgcan=@flagcancel,
			   cargoapecarabo=@cargoabono,
			   cargoaperefere=@referencia,
			   conceptocodigo=@concepto,
			   cargoapefecvct=@venci
                         Where  documentocargo=@tipodocu and cargonumdoc=@numero'
				
	SET @Parame = N'@tipodocu char(2),
			@numero  char(11),
			@cliente  char(11),
			@vendedor char(3),
			@zona 	char(3),
			@apefecemi datetime,
			@moneda	char(2),
			@apeimppag float,
			@usuario char(8),
			@tipocambio float,
			@fechaact  datetime,
                        @flagcancel  bit,
			@cargoabono char(1),
			@referencia varchar(200),
			@concepto char(2),
			@venci datetime'

	EXEC sp_executesql @cadena,@parame,@tipodocu, 
						@numero,
						@cliente,
						@vendedor,
						@zona,
						@apefecemi,
						@moneda,
						@apeimppag,
						@usuario,
						@tipocambio,
						@fechaact,
						@flagcancel,
						@cargoabono,
						@referencia,
						@concepto,
						@venci

 end     



















GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE FUNCTION TipoDoc(@tipo char(1),@importe float) 
RETURNS float
AS
BEGIN
DECLARE @monto float
SET @monto =
   CASE @tipo
	WHEN 'A' THEN @importe*-1
   	WHEN 'C' THEN @importe
   	ELSE @importe
   END
RETURN @monto
END



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

