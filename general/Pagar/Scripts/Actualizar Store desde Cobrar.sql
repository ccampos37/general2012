if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TipoDoc]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[TipoDoc]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[fn_datenumber]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[fn_datenumber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[fn_ubicacolumna]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[fn_ubicacolumna]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[fn_ubicarango]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[fn_ubicarango]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_CtaCteAbonosxCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_CtaCteAbonosxCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_CtaCtexCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_CtaCtexCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_DocumentosPendientes]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_DocumentosPendientes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_PlanillaCobranza]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_PlanillaCobranza]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_PlanillaDocVarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_PlanillaDocVarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_SaldoxCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_SaldoxCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_SaldoxCliente_Detalle]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_SaldoxCliente_Detalle]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_SubPlanillaCobranza]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_SubPlanillaCobranza]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_SubPlanillaDocVarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_SubPlanillaDocVarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_SubSaldoxCliente]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_SubSaldoxCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[cc_EMLB_SubSaldoxCliente_Detalle]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[cc_EMLB_SubSaldoxCliente_Detalle]
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

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE   function fn_datenumber(@dia int,@mes int,@año int) 
returns int
as
Begin
declare @fechavar datetime,@numfecha int
set @fechavar=cast(0 as datetime)

	
	Select @numfecha=Cast(	
	(dateadd(day,@dia- day(@fechavar),
	dateadd(month,@mes  - month(@fechavar),	
        dateadd(year,@año-(year(@fechavar)),cast(0 as datetime))))) as int)
return @numfecha
end

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE   function fn_ubicacolumna(@cad varchar(50),@numdia int )--,@ntotal int) 
returns int
as
Begin
declare @i as int
declare @num as int
set @i=1
while @i<=5
	begin
		set @num=left(@cad,patindex('%*%',@cad)-1)
		set @cad=right(@cad,len(@cad)-patindex('%*%',@cad))
		if @num>=@numdia 
			break

		set @i=@i+1
	end
return @i
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



--select dbo.fn_ubicarango('7*15*30*45*60*',5) 
--select 5

CREATE   function fn_ubicarango(@cad varchar(50),@ntope int )--,@ntotal int) 
returns int
as
Begin
declare @i as int
declare @num as int
set @i=1
while @i<=@ntope
	begin
		set @num=left(@cad,patindex('%*%',@cad)-1)
		set @cad=right(@cad,len(@cad)-patindex('%*%',@cad))
		set @i=@i+1
	end
return @num
end

/*
declare @fechavar datetime,@numfecha int
set @fechavar=cast(0 as datetime)

	
	Select @numfecha=Cast(	
	(dateadd(day,@dia- day(@fechavar),
	dateadd(month,@mes  - month(@fechavar),	
        dateadd(year,@año-(year(@fechavar)),cast(0 as datetime))))) as int)
return @numfecha
*/


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE  proc cc_EMLB_CtaCteAbonosxCliente(
@base varchar(50),
@compu varchar(20),
@fecha varchar(10),
@acuenta char(1), 
@codmoneda varchar(2),
@codcliente varchar(50),
@ctacontable varchar(20)
)
as

/*
declare @base varchar(50), @compu varchar(20), @fecha varchar(10), @acuenta char(1) 
declare @codmoneda varchar(2), @codcliente varchar(50), @ctacontable varchar(20)
set @base='ventas_prueba'
SET @compu='DESARROLLO3'
SET @fecha='18/10/2002'
SET @codmoneda='%'
SET @codcliente='%'
SET @ctacontable='%'
SET @acuenta='0'
*/

set nocount on
DECLARE @sqlcad varchar(3000)

declare @cadtmp varchar(2000)
set @cadtmp='SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
		B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
		B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
		B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla,
		simbmonabo=(select M.monedasimbolo from [' +@base+ '].dbo.gr_moneda as M where M.monedacodigo=B.abonocanmoncan),
		B.abonocantipcam
  		INTO ##tmp_saldodoc' +@compu+ '  
	FROM 
		[' +@base+ '].dbo.vt_abono B,
		[' +@base+ '].dbo.cc_tipodocumento C
	where abonocancli =''*'''

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldodoc'+@compu)
  exec('DROP TABLE ##tmp_saldodoc'+@compu)

if @acuenta='1'
BEGIN
  exec(@cadtmp)
END

if @acuenta='0'
BEGIN
	Set @sqlcad=' 
 		SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
			B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
			B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        	B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
			B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla,
			simbmonabo=(select M.monedasimbolo from [' +@base+ '].dbo.gr_moneda as M where M.monedacodigo=B.abonocanmoncan),
			B.abonocantipcam
  		INTO ##tmp_saldodoc' +@compu+ '  
		FROM 
			[' +@base+ '].dbo.vt_abono B,
			[' +@base+ '].dbo.cc_tipodocumento C,
 				(select clientecodigo,documentocargo,cargonumdoc from [' +@base+ '].dbo.vt_cargo D, [' +@base+ '].dbo.cc_tipodocumento E
 						where cargoapefecemi<=''' +@fecha+  '''  AND  cargoapeflgcan=0 AND
							cargoapeflgreg IS NULL AND
							D.documentocargo=E.tdocumentocodigo ) AS Z
 		WHERE 	B.abonocancli=Z.clientecodigo AND
        		B.documentoabono=Z.documentocargo AND 
				B.abononumdoc=Z.cargonumdoc AND
				B.abonocanfecan<='''  +@fecha+ ''' AND 
				B.abonocantdqc=C.tdocumentocodigo AND
				B.abonocancli like ''' +@codcliente+   ''''
  	exec(@sqlcad)

END

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldoactualizado'+@compu)
exec('DROP TABLE ##tmp_saldoactualizado' +@compu )

set @sqlcad='
	SELECT 	A.clientecodigo, A.documentocargo, A.cargonumdoc,A.cargoapefecemi,
			A.cargoapefecvct,A.bancocodigo,A.monedacodigo,cargoapeimpape=isnull(A.cargoapeimpape,0),
			cargoapeimppag=isnull(A.cargoapeimppag,0),
			cargopagadoux=isnull(A.cargoapeimppag,0),
			A.cargoapeflgcan,A.cargoapecarabo,
 			Y.*,
 			E.tdocumentodescripcion,G.bancodescripcion,
			I.clienteruc,I.clienterazonsocial,
			H.monedasimbolo
    INTO 	##tmp_saldoactualizado' +@compu+ '
	FROM 	[' +@base+ '].dbo.vt_cargo A, ##tmp_saldodoc' +@compu+ ' Y,
 			[' +@base+ '].dbo.cc_tipodocumento E,
			[' +@base+ '].dbo.gr_banco G,
			[' +@base+ '].dbo.gr_moneda H,
			[' +@base+ '].dbo.vt_cliente I
	WHERE 	A.clientecodigo*=Y.abonocancli AND
        	A.documentocargo*=Y.documentoabono AND 
			A.cargonumdoc*=Y.abononumdoc AND
	    	A.documentocargo=E.tdocumentocodigo AND
			A.bancocodigo*=G.bancocodigo AND
			A.monedacodigo=H.monedacodigo AND
        	A.clientecodigo=I.clientecodigo AND
			A.clientecodigo like ''' +@codcliente+ ''' AND 
			A.monedacodigo like ''' +@codmoneda+ '''	AND
            A.cargoapefecemi<=''' +@fecha+ ''' AND
			A.cargoapeflgcan=0  AND
			A.cargoapeflgreg IS NULL 
	ORDER BY cast(A.clientecodigo as int), A.documentocargo,A.cargonumdoc'

exec(@sqlcad)


IF @acuenta='0' 
BEGIN
	set @sqlcad='
		UPDATE ##tmp_saldoactualizado' +@compu+ ' SET ##tmp_saldoactualizado' +@compu+ '.cargopagadoux=B.saldoactual
			FROM ##tmp_saldoactualizado' +@compu+ ' A,
				 (SELECT clientecodigo,documentocargo,cargonumdoc,saldoactual=SUM(ISNULL(abonocanimpsol,0)) 
					FROM ##tmp_saldoactualizado' +@compu+ ' 
 					GROUP BY clientecodigo,documentocargo,cargonumdoc ) as B
			WHERE A.clientecodigo=B.clientecodigo AND A.documentocargo=B.documentocargo AND
	  			A.cargonumdoc=B.cargonumdoc'

	exec(@sqlcad)
END

IF @fecha<convert(varchar(10),getdate(),103) AND @acuenta='1'
BEGIN
	exec('UPDATE ##tmp_saldoactualizado' +@compu+ ' SET cargopagadoux=0')
		
	SET @sqlcad=''
	SET @sqlcad='
	UPDATE ##tmp_saldoactualizado' +@compu+ ' SET ##tmp_saldoactualizado' +@compu+ '.cargopagadoux=Y.saldoactual
				FROM ##tmp_saldoactualizado' +@compu+ ' A,
			(SELECT B.abonocancli,B.documentoabono,B.abononumdoc,saldoactual=SUM(ISNULL(B.abonocanimpsol,0))
			FROM 
				[' +@base+ '].dbo.vt_abono B,
				[' +@base+ '].dbo.cc_tipodocumento C,
 					(select clientecodigo,documentocargo,cargonumdoc from [' +@base+ '].dbo.vt_cargo D, [' +@base+ '].dbo.cc_tipodocumento E
	 							where cargoapefecemi<=''' +@fecha+  '''  AND  cargoapeflgcan=0 AND
								cargoapeflgreg IS NULL AND
								D.documentocargo=E.tdocumentocodigo ) AS Z
	 		WHERE 	B.abonocancli=Z.clientecodigo AND
    	    		B.documentoabono=Z.documentocargo AND 
					B.abononumdoc=Z.cargonumdoc AND
					B.abonocanfecan<='''  +@fecha+ ''' AND 
					B.abonocantdqc=C.tdocumentocodigo AND
					B.abonocancli like ''' +@codcliente+   '''
        	GROUP BY B.abonocancli,B.documentoabono,B.abononumdoc) as Y
		WHERE A.clientecodigo=Y.abonocancli AND A.documentocargo=Y.documentoabono AND
	  		A.cargonumdoc=Y.abononumdoc'
	exec(@sqlcad)

END

exec('SELECT * FROM  ##tmp_saldoactualizado' +@compu )

set nocount off

--select * from ##tmp_saldodocdesarrollo3 order by abonocancli,documentoabono,abononumdoc
--exec cc_EMLB_SaldoxCliente_Detalle 'ventas_prueba','DESARROLLO3','18/11/2002','1','%','%','%'





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO









CREATE       proc cc_EMLB_CtaCtexCliente(
@base varchar(50),
@compu varchar(20),
@fechaini varchar(10),
@fechafin varchar(10),
@fecha varchar(10),
@codmoneda varchar(2),
@codcliente varchar(50)
)
as

set nocount on
DECLARE @sqlcad varchar(3000)
declare @cadtmp varchar(2000)
set @cadtmp='SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
		B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
		B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
		B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla,
		simbmonabo=(select M.monedasimbolo from [' +@base+ '].dbo.gr_moneda as M where M.monedacodigo=B.abonocanmoncan),
		B.abonocantipcam
  		INTO ##tmp_saldodoc' +@compu+ '  
	FROM 
		[' +@base+ '].dbo.vt_abono B,
		[' +@base+ '].dbo.cc_tipodocumento C
	where abonocancli =''*'''

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldodoc'+@compu)
  exec('DROP TABLE ##tmp_saldodoc'+@compu)

exec(@cadtmp)

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldoinicial'+@compu)
exec('DROP TABLE ##tmp_saldoinicial' +@compu )

set @sqlcad='
	SELECT 	A.clientecodigo, A.documentocargo, A.cargonumdoc,A.cargoapefecemi,
			A.cargoapefecvct,A.bancocodigo,A.monedacodigo,cargoapeimpape=isnull(A.cargoapeimpape,0),
			cargoapeimppag=isnull(A.cargoapeimppag,0),
			cargopagadoux=isnull(A.cargoapeimppag,0),
			SaldoInicial=cast(0 as numeric(25,9)),
			A.cargoapeflgcan,A.cargoapecarabo,
 			Y.*,
 			E.tdocumentodescripcion,G.bancodescripcion,
			I.clienteruc,I.clienterazonsocial,
			H.monedasimbolo
    INTO 	##tmp_saldoinicial' +@compu+ '
	FROM 	[' +@base+ '].dbo.vt_cargo A, ##tmp_saldodoc' +@compu+ ' Y,
 			[' +@base+ '].dbo.cc_tipodocumento E,
			[' +@base+ '].dbo.gr_banco G,
			[' +@base+ '].dbo.gr_moneda H,
			[' +@base+ '].dbo.vt_cliente I
	WHERE 	A.clientecodigo*=Y.abonocancli AND
        	A.documentocargo*=Y.documentoabono AND 
			A.cargonumdoc*=Y.abononumdoc AND
	    	A.documentocargo=E.tdocumentocodigo AND
			A.bancocodigo*=G.bancocodigo AND
			A.monedacodigo=H.monedacodigo AND
        	A.clientecodigo=I.clientecodigo AND
			A.clientecodigo like ''' +@codcliente+ ''' AND 
			A.monedacodigo like ''' +@codmoneda+ '''	AND
            floor(cast(A.cargoapefecemi as real)) <=' + cast(dbo.fn_datenumber(day(@fecha),month(@fecha),year(@fecha)) as varchar(20))  + ' AND
			A.cargoapeflgcan=0  AND
			A.cargoapeflgreg IS NULL
	ORDER BY cast(A.clientecodigo as int), A.documentocargo,A.cargonumdoc'

exec(@sqlcad)

--IF @fecha<convert(varchar(10),getdate(),103)
--BEGIN
	exec('UPDATE ##tmp_saldoinicial' +@compu+ ' SET cargopagadoux=0')
	SET @sqlcad=''
	SET @sqlcad=N'
	UPDATE ##tmp_saldoinicial' +@compu+ ' SET ##tmp_saldoinicial' +@compu+ '.cargopagadoux=Y.saldoactual
				FROM ##tmp_saldoinicial' +@compu+ ' A,
			(SELECT B.abonocancli,B.documentoabono,B.abononumdoc,saldoactual=SUM(ISNULL(B.abonocanimpsol,0))
			FROM 
				[' +@base+ '].dbo.vt_abono B,
				[' +@base+ '].dbo.cc_tipodocumento C,
 					(select clientecodigo,documentocargo,cargonumdoc from [' +@base+ '].dbo.vt_cargo D, [' +@base+ '].dbo.cc_tipodocumento E
	 							where FLOOR(CAST(cargoapefecemi AS REAL)) <=' + CAST(DBO.fn_datenumber(DAY(@fecha),MONTH(@fecha),YEAR(@fecha)) as varchar(20)) +  '  AND  cargoapeflgcan=0 AND
								cargoapeflgreg IS NULL AND
								D.documentocargo=E.tdocumentocodigo ) AS Z
	 		WHERE 	B.abonocancli=Z.clientecodigo AND
    	    		B.documentoabono=Z.documentocargo AND 
					B.abononumdoc=Z.cargonumdoc AND
					FLOOR(CAST (B.abonocanfecan AS REAL)) <='  + CAST(DBO.fn_datenumber(DAY(@fecha),MONTH(@FECHA),YEAR(@FECHA)) AS VARCHAR(20)) + ' AND 
					B.abonocantdqc=C.tdocumentocodigo AND
					B.abonocancli like ''' +@codcliente+   '''
        	GROUP BY B.abonocancli,B.documentoabono,B.abononumdoc) as Y
		WHERE A.clientecodigo=Y.abonocancli AND A.documentocargo=Y.documentoabono AND
	  		A.cargonumdoc=Y.abononumdoc'
	exec(@sqlcad)
--END
exec('UPDATE ##tmp_saldoinicial' +@compu+ ' SET saldoinicial=(cargoapeimpape-cargopagadoux)')

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_abonos'+@compu)
exec('DROP TABLE ##tmp_abonos' +@compu )

declare @cadsql nvarchar(4000)
SET @cadsql='
	SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
			B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
			B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        	B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
			B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla,
			simbmonabo=D.monedasimbolo,
			B.abonocantipcam
		INTO ##tmp_abonos' +@compu+ '
		FROM 
			[' +@base+ '].dbo.vt_abono B,
			[' +@base+ '].dbo.cc_tipodocumento C,
			[' +@base+ '].dbo.gr_moneda D
 		WHERE 	
			B.abonocanfecan between ''' +@fechaini+ ''' AND ''' +@fechafin+ ''' AND
			B.abonocantdqc=C.tdocumentocodigo AND
			B.abonocanmoncan=D.monedacodigo AND 
			B.abonocancli LIKE ''' +@codcliente+   '''' 
exec(@cadsql)

SET @cadsql='
	SELECT 	AA.*,
			C.clienteruc,C.clienterazonsocial,
			descdoccargo=D.tdocumentodescripcion,
			numplanillacargo=E.abononumplanilla
		FROM	
			[' +@base+ '].dbo.vt_cliente C,
			[' +@base+ '].dbo.cc_tipodocumento D,
			[' +@base+ '].dbo.vt_cargo E,
			(SELECT A.clientecodigo, A.documentocargo, A.cargonumdoc, B.abononumplanilla, A.cargoapefecemi,A.cargoapefecvct, A.monedacodigo, A.monedasimbolo, A.cargoapeimpape,
				   	B.abonocantdqc,B.tipodescripcion,B.abonocanndqc,B.abonocanmoncan,abonocanimpcan=ISNULL(B.abonocanimpcan,0),abonocanimpsol=ISNULL(B.abonocanimpsol,0),
			   		B.abonocanfecan, B.simbmonabo,
		  			SaldoInicial
			FROM  
				##tmp_saldoinicial' +@compu+ ' A, ##tmp_abonos' +@compu+ ' B
   			WHERE 
				A.clientecodigo=B.abonocancli AND
	    		A.documentocargo=B.documentoabono AND
				A.cargonumdoc=B.abononumdoc AND
				B.abonocancli LIKE ''' +@codcliente+  '''
			UNION ALL	
			SELECT 	A.clientecodigo,A.documentocargo,A.cargonumdoc, B.abononumplanilla, A.cargoapefecemi, A.cargoapefecvct, A.monedacodigo,C.monedasimbolo, A.cargoapeimpape, 
					B.abonocantdqc,B.tipodescripcion,B.abonocanndqc,B.abonocanmoncan,abonocanimpcan=ISNULL(B.abonocanimpcan,0),abonocanimpsol=ISNULL(B.abonocanimpsol,0) ,
					B.abonocanfecan, B.simbmonabo,
					SaldoInicial=cast(0 as numeric(25,9))
			FROM  
				[' +@base+ '].dbo.vt_cargo A, ##tmp_abonos' +@compu+ ' B, [' +@base+ '].dbo.gr_moneda C
			WHERE 
				A.clientecodigo*=B.abonocancli AND
	    		A.documentocargo*=B.documentoabono AND
				A.cargonumdoc*=B.abononumdoc AND
				A.cargoapefecemi between ''' +@fechaini+ ''' AND ''' +@fechafin+ ''' AND
                B.abonocancli LIKE ''' +@codcliente+ ''' AND A.monedacodigo=C.monedacodigo) AS AA
		WHERE 
			C.clientecodigo=AA.clientecodigo AND
			D.tdocumentocodigo=AA.documentocargo AND
			E.clientecodigo=AA.clientecodigo AND
			E.documentocargo=AA.documentocargo AND
			E.cargonumdoc=AA.cargonumdoc AND
			AA.clientecodigo LIKE ''' +@codcliente+ ''' AND
			AA.monedacodigo  LIKE ''' +@codmoneda+ ''''
			
exec(@cadsql)

set nocount off

--select * from ##tmp_saldodocdesarrollo3 order by abonocancli,documentoabono,abononumdoc
--exec cc_EMLB_CtaCtexCliente 'ventas_prueba','DESARROLLO3','15/11/2002','30/11/2002','14/11/2002','%','338'






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





/*Documentos Vencidos y por Vencer*/
CREATE     proc cc_EMLB_DocumentosPendientes(
@base varchar(50),
@compu varchar(20),
@fecha varchar(10),
@RangoVencer varchar(50),
@RangoVencido varchar(50),
@codcliente varchar(20),
@codmoneda varchar(2)
)
as

set nocount on
DECLARE @sqlcad varchar(3500)
DECLARE @totdiaven as integer
DECLARE @totdiapve as integer
declare @valortope as integer


set @valortope=dbo.fn_ubicarango('' +@RangoVencer+ '',5)

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_documpendiente'+@compu)
exec('DROP TABLE ##tmp_documpendiente' +@compu )

set @sqlcad='
	SELECT 	A.clientecodigo, A.documentocargo, A.cargonumdoc,A.cargoapefecemi,
			A.cargoapefecvct,A.monedacodigo,cargoapeimpape=isnull(A.cargoapeimpape,0),
			cargoapeimppag=isnull(A.cargoapeimppag,0),
			cargopagadoux=isnull(A.cargoapeimppag,0),
			A.cargoapeflgcan,A.cargoapecarabo,
 			E.tdocumentodescripcion,
			I.clienteruc,I.clienterazonsocial,
			H.monedasimbolo,
			estadoreg=case when floor(cast(cargoapefecvct-''' +@fecha+ ''' as real))>=0
				then ''POR VENCER''
				else ''VENCIDO''
			end,
			numdias=floor(cast(cargoapefecvct-''' +@fecha+ ''' as real)),
			numcolumna=dbo.fn_ubicacolumna(''' +@RangoVencer+ ''',abs(floor(cast(cargoapefecvct-''' +@fecha+ ''' as real)))),
			DocVen1=cast(0 as numeric(25,9)),DocVen2=cast(0 as numeric(25,9)),DocVen3=cast(0 as numeric(25,9)),DocVen4=cast(0 as numeric(25,9)),DocVen5=cast(0 as numeric(25,9)),
			DocPVe1=cast(0 as numeric(25,9)),DocPVe2=cast(0 as numeric(25,9)),DocPVe3=cast(0 as numeric(25,9)),DocPVe4=cast(0 as numeric(25,9)),DocPVe5=cast(0 as numeric(25,9))
    INTO 	##tmp_documpendiente' +@compu+ '
	FROM 	[' +@base+ '].dbo.vt_cargo A,
 			[' +@base+ '].dbo.cc_tipodocumento E,
			[' +@base+ '].dbo.gr_moneda H,
			[' +@base+ '].dbo.vt_cliente I
	WHERE 	A.documentocargo=E.tdocumentocodigo AND
			A.monedacodigo=H.monedacodigo AND
        	A.clientecodigo=I.clientecodigo AND
			A.clientecodigo like ''' +@codcliente+ ''' AND 
			A.monedacodigo like ''' +@codmoneda+ '''	AND
			A.cargoapeflgcan=0  AND
			A.cargoapeflgreg IS NULL AND
			cast(abs(floor(cast(A.cargoapefecvct-''' +@fecha+ ''' as real))) as integer)<=' +cast(@valortope as varchar(2))+  ' 
	ORDER BY cast(A.clientecodigo as int), A.documentocargo,A.cargonumdoc'

exec(@sqlcad)

declare @cadsql varchar(3000)
declare @i as integer

set @cadsql='update ##tmp_documpendiente' +@compu+ ' set DocVen1=(cargoapeimpape-cargoapeimppag)
             where numcolumna=0 AND estadoreg=''POR VENCER'''
EXEC(@cadsql) 

set @i=1
while @i<=5
begin
	set @cadsql='update ##tmp_documpendiente' +@compu+ ' set DocVen' +cast(@i as varchar(2)) +  '=(cargoapeimpape-cargoapeimppag)
             where numcolumna=' +cast(@i as varchar(2)) + ' AND estadoreg=''VENCIDO'''
	EXEC(@cadsql) 
	set @i=@i+1
end

set @i=1
while @i<=5
begin
	set @cadsql='update ##tmp_documpendiente' +@compu+ ' set DocPVe' +cast(@i as varchar(2)) +  '=(cargoapeimpape-cargoapeimppag)
             where numcolumna=' +cast(@i as varchar(2)) + ' AND estadoreg=''POR VENCER'''
	EXEC(@cadsql) 
	set @i=@i+1
end

exec('SELECT * FROM ##tmp_documpendiente' +@compu)

set nocount off

--select * from ##tmp_saldodocdesarrollo3 order by abonocancli,documentoabono,abononumdoc
--exec cc_EMLB_DocumentosPendientes 'ventas_prueba','DESARR','31/10/2002','7*15*30*45*60*','7*15*30*45*60*','%','%'






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO






CREATE    PROC cc_EMLB_PlanillaCobranza 		/*EN USO*/
@base varchar(50),
@fecdesde varchar(10),
@fechasta varchar(10),
@codvendedor nvarchar(3)

AS

DECLARE @sensql nvarchar (4000)
SET @sensql = N'
SELECT 	e.abonocancli as Cod_Cliente,c.clienterazonsocial as Razon_Social,
	e.vendedorcodigo as Cod_Vendedor,e.documentoabono as Cod_Doc_Cargo,
	b.tdocumentodesccorta as Desc_Doc_Cargo,e.abononumdoc as Num_Doc_Cargo,
	a.cargoapefecemi as Fec_Emision_Cargo,
	e.abonocanforcan as Forma_Pago,d.monedasimbolo, 
	ISNULL(e.abonocanimpcan,0) as Importe_Abono,
    e.abonocanmoncan as MonedaAbono,
    simbolo_mon_abo=(select M.monedasimbolo from [' +@base+ '].dbo.gr_moneda as M where M.monedacodigo=e.abonocanmoncan),
	e.abonocantdqc as Cod_Doc_Abono,g.tdocumentodesccorta as Desc_Doc_Abono,
	e.abonocanndqc as Num_Doc_Abono,e.abonocanfecan as Fec_Cancela_Abono,
	e.abononumplanilla as Num_Planilla, e.abonotipoplanilla as Tipo_Planilla,
	tipopago=
		case when a.cargoapefecemi=e.abonocanfecan and e.abonocanforcan=''T''
			then ''CO'' else ''CR''
		end
FROM 	
	['+@base+'].dbo.vt_cargo a 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento b 
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN
	['+@base+'].dbo.vt_cliente c
	ON
	a.clientecodigo = c.clientecodigo
	JOIN
	['+@base+'].dbo.gr_moneda d
	ON
	a.monedacodigo = d.monedacodigo
	JOIN
	['+@base+'].dbo.vt_abono e
	ON
	a.documentocargo = e.documentoabono
	AND a.cargonumdoc = e.abononumdoc
	JOIN
	['+@base+'].dbo.cc_tipodocumento g
	ON
	g.tdocumentocodigo = e.abonocantdqc
	JOIN
	['+@base+'].dbo.cc_tipoplanilla h
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



CREATE      PROC cc_EMLB_PlanillaDocVarios 		/*EN USO*/
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
	['+@base+'].dbo.vt_cargo a 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento b 
	ON 
	a.documentocargo = b.tdocumentocodigo 
	JOIN 
	['+@base+'].dbo.vt_cliente c 
	ON 
	a.clientecodigo = c.clientecodigo 
	JOIN 
	['+@base+'].dbo.gr_moneda d 
	ON 
	a.monedacodigo = d.monedacodigo 
	JOIN 
	['+@base+'].dbo.cc_tipoplanilla f 
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


CREATE       PROC cc_EMLB_SaldoxCliente 		/*EN USO*/
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
	['+@base+'].dbo.vt_cargo a 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento b 
	ON 
	a.documentocargo = b.tdocumentocodigo 
	JOIN 
	['+@base+'].dbo.vt_cliente c 
	ON 
	a.clientecodigo = c.clientecodigo 
	JOIN 
	['+@base+'].dbo.gr_moneda d 
	ON 
	a.monedacodigo = d.monedacodigo 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento f 
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
	['+@base+'].dbo.vt_cargo a 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento b
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN
	['+@base+'].dbo.vt_cliente c
	ON
	a.clientecodigo = c.clientecodigo
	JOIN
	['+@base+'].dbo.gr_moneda d
	ON
	a.monedacodigo = d.monedacodigo
	JOIN
	['+@base+'].dbo.vt_abono e
	ON
	a.documentocargo = e.documentoabono
	AND a.cargonumdoc = e.abononumdoc
	JOIN
	['+@base+'].dbo.cc_tipodocumento f
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







CREATE       proc cc_EMLB_SaldoxCliente_Detalle(
@base varchar(50),
@compu varchar(20),
@fecha varchar(10),
@acuenta char(1), 
@codmoneda varchar(2),
@codcliente varchar(50),
@ctacontable varchar(20)
)
as

/*
declare @base varchar(50), @compu varchar(20), @fecha varchar(10), @acuenta char(1) 
declare @codmoneda varchar(2), @codcliente varchar(50), @ctacontable varchar(20)
set @base='ventas_prueba'
SET @compu='DESARROLLO3'
SET @fecha='18/10/2002'
SET @codmoneda='%'
SET @codcliente='%'
SET @ctacontable='%'
SET @acuenta='0'
*/

set nocount on
DECLARE @sqlcad varchar(3000)
DECLARE @condctacontable nvarchar (2000)
SET @condctacontable=''
IF @ctacontable <> '%'
BEGIN
  SET @condctacontable=' AND(LTRIM(RTRIM(E.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') OR LTRIM(RTRIM(E.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''')) '
  IF @codmoneda <> '%'
    BEGIN
	  IF @codmoneda = '01'
	   BEGIN
		 SET @condctacontable=' AND LTRIM(RTRIM(E.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') '
	   END
	  IF @codmoneda = '02'
	   BEGIN
		 SET @condctacontable= ' AND LTRIM(RTRIM(E.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''') '
	   END
	END
END

declare @cadtmp varchar(2000)
set @cadtmp='SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
		B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
		B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
		B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla,
		simbmonabo=(select M.monedasimbolo from [' +@base+ '].dbo.gr_moneda as M where M.monedacodigo=B.abonocanmoncan),
		B.abonocantipcam
  		INTO ##tmp_saldodoc' +@compu+ '  
	FROM 
		[' +@base+ '].dbo.vt_abono B,
		[' +@base+ '].dbo.cc_tipodocumento C
	where abonocancli =''*'''

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldodoc'+@compu)
  exec('DROP TABLE ##tmp_saldodoc'+@compu)

if @acuenta='1'
BEGIN
  exec(@cadtmp)
END

if @acuenta='0'
BEGIN
	Set @sqlcad=' 
 		SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
			B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
			B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        	B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
			B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla,
			simbmonabo=(select M.monedasimbolo from [' +@base+ '].dbo.gr_moneda as M where M.monedacodigo=B.abonocanmoncan),
			B.abonocantipcam
  		INTO ##tmp_saldodoc' +@compu+ '  
		FROM 
			[' +@base+ '].dbo.vt_abono B,
			[' +@base+ '].dbo.cc_tipodocumento C,
 				(select clientecodigo,documentocargo,cargonumdoc from [' +@base+ '].dbo.vt_cargo D, [' +@base+ '].dbo.cc_tipodocumento E
 						where cargoapefecemi<=''' +@fecha+  '''  AND  cargoapeflgcan=0 AND
							cargoapeflgreg IS NULL ' +@condctacontable+ ' AND
							D.documentocargo=E.tdocumentocodigo ) AS Z
 		WHERE 	B.abonocancli=Z.clientecodigo AND
        		B.documentoabono=Z.documentocargo AND 
				B.abononumdoc=Z.cargonumdoc AND
				B.abonocanfecan<='''  +@fecha+ ''' AND 
				B.abonocantdqc=C.tdocumentocodigo AND
				B.abonocancli like ''' +@codcliente+   ''''
  	exec(@sqlcad)

END

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldoactualizado'+@compu)
exec('DROP TABLE ##tmp_saldoactualizado' +@compu )

set @sqlcad='
	SELECT 	A.clientecodigo, A.documentocargo, A.cargonumdoc,A.cargoapefecemi,
			A.cargoapefecvct,A.bancocodigo,A.monedacodigo,
			cargoapeimpape=ISNULL( dbo.tipodoc(E.tdocumentotipo,A.cargoapeimpape) ,0 ),
			cargoapeimppag=ISNULL(A.cargoapeimppag,0 ),
			cargopagadoux=isnull(A.cargoapeimppag,0),
			A.cargoapeflgcan,A.cargoapecarabo,
 			Y.*,
 			E.tdocumentodescripcion,G.bancodescripcion,
			I.clienteruc,I.clienterazonsocial,
			H.monedasimbolo
    INTO 	##tmp_saldoactualizado' +@compu+ '
	FROM 	[' +@base+ '].dbo.vt_cargo A, ##tmp_saldodoc' +@compu+ ' Y,
 			[' +@base+ '].dbo.cc_tipodocumento E,
			[' +@base+ '].dbo.gr_banco G,
			[' +@base+ '].dbo.gr_moneda H,
			[' +@base+ '].dbo.vt_cliente I
	WHERE 	A.clientecodigo*=Y.abonocancli AND
        	A.documentocargo*=Y.documentoabono AND 
			A.cargonumdoc*=Y.abononumdoc AND
	    	A.documentocargo=E.tdocumentocodigo AND
			A.bancocodigo*=G.bancocodigo AND
			A.monedacodigo=H.monedacodigo AND
        	A.clientecodigo=I.clientecodigo AND
			A.clientecodigo like ''' +@codcliente+ ''' AND 
			A.monedacodigo like ''' +@codmoneda+ '''	AND
            A.cargoapefecemi<=''' +@fecha+ ''' AND
			A.cargoapeflgcan=0  AND
			A.cargoapeflgreg IS NULL ' +@condctacontable + '
	ORDER BY cast(A.clientecodigo as int), A.documentocargo,A.cargonumdoc'

exec(@sqlcad)


IF @acuenta='0' 
BEGIN
	set @sqlcad='
		UPDATE ##tmp_saldoactualizado' +@compu+ ' SET ##tmp_saldoactualizado' +@compu+ '.cargopagadoux=B.saldoactual
			FROM ##tmp_saldoactualizado' +@compu+ ' A,
				 (SELECT clientecodigo,documentocargo,cargonumdoc,saldoactual=SUM(ISNULL(abonocanimpsol,0)) 
					FROM ##tmp_saldoactualizado' +@compu+ ' 
 					GROUP BY clientecodigo,documentocargo,cargonumdoc ) as B
			WHERE A.clientecodigo=B.clientecodigo AND A.documentocargo=B.documentocargo AND
	  			A.cargonumdoc=B.cargonumdoc'

	exec(@sqlcad)
END

IF @fecha<convert(varchar(10),getdate(),103) AND @acuenta='1'
BEGIN
	exec('UPDATE ##tmp_saldoactualizado' +@compu+ ' SET cargopagadoux=0')
		
	SET @sqlcad=''
	SET @sqlcad='
	UPDATE ##tmp_saldoactualizado' +@compu+ ' SET ##tmp_saldoactualizado' +@compu+ '.cargopagadoux=Y.saldoactual
				FROM ##tmp_saldoactualizado' +@compu+ ' A,
			(SELECT B.abonocancli,B.documentoabono,B.abononumdoc,saldoactual=SUM(ISNULL(B.abonocanimpsol,0))
			FROM 
				[' +@base+ '].dbo.vt_abono B,
				[' +@base+ '].dbo.cc_tipodocumento C,
 					(select clientecodigo,documentocargo,cargonumdoc from [' +@base+ '].dbo.vt_cargo D, [' +@base+ '].dbo.cc_tipodocumento E
	 							where cargoapefecemi<=''' +@fecha+  '''  AND  cargoapeflgcan=0 AND
								cargoapeflgreg IS NULL ' +@condctacontable+ ' AND
								D.documentocargo=E.tdocumentocodigo ) AS Z
	 		WHERE 	B.abonocancli=Z.clientecodigo AND
    	    		B.documentoabono=Z.documentocargo AND 
					B.abononumdoc=Z.cargonumdoc AND
					B.abonocanfecan<='''  +@fecha+ ''' AND 
					B.abonocantdqc=C.tdocumentocodigo AND
					B.abonocancli like ''' +@codcliente+   '''
        	GROUP BY B.abonocancli,B.documentoabono,B.abononumdoc) as Y
		WHERE A.clientecodigo=Y.abonocancli AND A.documentocargo=Y.documentoabono AND
	  		A.cargonumdoc=Y.abononumdoc'
	exec(@sqlcad)

END

exec('SELECT * FROM  ##tmp_saldoactualizado' +@compu )

set nocount off

--select * from ##tmp_saldodocdesarrollo3 order by abonocancli,documentoabono,abononumdoc
--exec cc_EMLB_SaldoxCliente_Detalle 'ventas_prueba','DESARROLLO3','18/11/2002','1','%','124','%'








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO





CREATE   proc cc_EMLB_SubPlanillaCobranza (
@base varchar(50),
@fecdesde varchar(10),
@fechasta varchar(10),
@codvendedor nvarchar(3)
)
as
/*
set @base='ventas_prueba'
set @fecdesde='01/11/2002'
set @fechasta='25/11/2002'
set @codvendedor='%'
*/

DECLARE @sensql1 nvarchar (4000)
DECLARE @sensql2 nvarchar (4000)
SET @sensql1= '

SELECT Cod_Doc_Abono,Desc_Doc_Abono,Cod_Doc_Cargo,Desc_Doc_Cargo,
	ContadoDolar=IMPORTES_DOLARES_CONTADO,
    CreditoDolar=(IMPORTES_DOLARES_TOTAL-IMPORTES_DOLARES_CONTADO),
	ContadoSol=IMPORTES_SOLES_CONTADO,
	CreditoSol=(IMPORTES_SOLES_TOTAL-IMPORTES_SOLES_CONTADO)
 FROM

(SELECT e.abonocantdqc as Cod_Doc_Abono,g.tdocumentodescripcion as Desc_Doc_Abono,
	e.documentoabono as Cod_Doc_Cargo,i.tdocumentodescripcion as Desc_Doc_Cargo,
	IMPORTES_DOLARES_CONTADO = 
	isnull ( (
	SELECT SUM (isnull(z.abonocanimpcan,0)) 
	FROM [' +@base+ '].dbo.vt_abono z
	JOIN
	[' +@base+ '].dbo.cc_tipoplanilla y
	ON
	y.tplanillacodigo = z.abonotipoplanilla
	JOIN
		[' +@base+ '].dbo.vt_cargo c
    ON
		c.documentocargo=z.documentoabono and  c.cargonumdoc=z.abononumdoc and c.clientecodigo=z.abonocancli 

	WHERE 	z.abonocanforcan=''T'' AND c.cargoapefecemi=z.abonocanfecan
	AND	z.abonocanmoncan = ''02'' 
	AND z.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND z.abonocanflreg IS NULL
	AND z.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND y.tplanillacobranza = ''1'' 
	AND e.abonocantdqc = z.abonocantdqc AND e.documentoabono = z.documentoabono
	) , 0 ) ,

	IMPORTES_DOLARES_TOTAL = 
	isnull ( (
	SELECT SUM (isnull(z.abonocanimpcan,0)) 
	FROM [' +@base+ '].dbo.vt_abono z
	JOIN
		[' +@base+ '].dbo.cc_tipoplanilla y
	ON
		y.tplanillacodigo = z.abonotipoplanilla
	WHERE
	z.abonocanmoncan = ''02'' 
	AND z.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND z.abonocanflreg IS NULL
	AND z.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND y.tplanillacobranza = ''1'' 
	AND e.abonocantdqc = z.abonocantdqc AND e.documentoabono = z.documentoabono
	) , 0 ) , '

set @sensql2='
	IMPORTES_SOLES_CONTADO = 
	isnull ( (
	SELECT SUM (isnull(z.abonocanimpcan,0)) 
	FROM [' +@base+ '].dbo.vt_abono z
	JOIN
		[' +@base+ '].dbo.cc_tipoplanilla y
	ON
		y.tplanillacodigo = z.abonotipoplanilla
	JOIN
		[' +@base+ '].dbo.vt_cargo c
    ON
		c.documentocargo=z.documentoabono and  c.cargonumdoc=z.abononumdoc and c.clientecodigo=z.abonocancli 

	WHERE 	z.abonocanforcan=''T'' AND c.cargoapefecemi=z.abonocanfecan
	AND	z.abonocanmoncan = ''01'' 
	AND z.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND z.abonocanflreg IS NULL
	AND z.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND y.tplanillacobranza = ''1'' 
	AND e.abonocantdqc = z.abonocantdqc AND e.documentoabono = z.documentoabono
	) , 0 ),

	IMPORTES_SOLES_TOTAL = 
	isnull ( (
	SELECT SUM (isnull(z.abonocanimpcan,0)) 
	FROM [' +@base+ '].dbo.vt_abono z
	JOIN
		[' +@base+ '].dbo.cc_tipoplanilla y
	ON
		y.tplanillacodigo = z.abonotipoplanilla
	WHERE
	z.abonocanmoncan = ''01'' 
	AND z.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND z.abonocanflreg IS NULL
	AND z.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND y.tplanillacobranza = ''1'' 
	AND e.abonocantdqc = z.abonocantdqc AND e.documentoabono = z.documentoabono
	) , 0 )

FROM 	
		[' +@base+ '].dbo.vt_abono e
	JOIN
		[' +@base+ '].dbo.cc_tipodocumento g
	ON
		g.tdocumentocodigo = e.abonocantdqc
	JOIN
		[' +@base+ '].dbo.cc_tipoplanilla h
	ON
		h.tplanillacodigo = e.abonotipoplanilla
	JOIN
		[' +@base+ '].dbo.cc_tipodocumento i
	ON
		i.tdocumentocodigo = e.documentoabono
WHERE	
	e.abonocanfecpla BETWEEN '''+@fecdesde+''' AND '''+@fechasta+''' 
	AND e.vendedorcodigo LIKE '''+@codvendedor+''' 
	AND h.tplanillacobranza = ''1''
	AND e.abonocanflreg IS NULL
GROUP BY 
	e.abonocantdqc,g.tdocumentodescripcion,
	e.documentoabono,i.tdocumentodescripcion ) AS ZZ'

exec (@sensql1+@sensql2)





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE    PROC cc_EMLB_SubPlanillaDocVarios 		/*EN USO*/
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
	['+@base+'].dbo.vt_cargo a 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento b 
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN 
	['+@base+'].dbo.cc_tipoplanilla f 
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


CREATE proc cc_EMLB_SubSaldoxCliente 		/*EN USO*/
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
	['+@base+'].dbo.vt_cargo a 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento b 
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
--	['+@base+'].dbo.cc_tipodocumento f 
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
	['+@base+'].dbo.vt_cargo a 
	JOIN 
	['+@base+'].dbo.cc_tipodocumento b
	ON  
	a.documentocargo = b.tdocumentocodigo 
	JOIN
	['+@base+'].dbo.vt_abono e
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
--	['+@base+'].dbo.cc_tipodocumento f
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







CREATE       proc cc_EMLB_SubSaldoxCliente_Detalle(
@base varchar(50),
@compu varchar(20),
@fecha varchar(10),
@acuenta char(1), 
@codmoneda varchar(2),
@codcliente varchar(50),
@ctacontable varchar(20)
/*@codresumen char(1)*/
)
as

/*
declare @base varchar(50), @compu varchar(20), @fecha varchar(10), @acuenta char(1) 
declare @codmoneda varchar(2), @codcliente varchar(50), @ctacontable varchar(20)
set @base='ventas_prueba'
SET @compu='DESARROLLO3'
SET @fecha='30/09/2002'
SET @codmoneda='%'
SET @codcliente='%'
SET @ctacontable='%'
SET @acuenta='1'
*/

set nocount on
DECLARE @sqlcad varchar(3000)
DECLARE @condctacontable nvarchar (2000)

--IF @CODRESUMEN='1'
--BEGIN

SET @condctacontable=''
IF @ctacontable <> '%'
BEGIN
  SET @condctacontable=' AND(LTRIM(RTRIM(E.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') OR LTRIM(RTRIM(E.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''')) '
  IF @codmoneda <> '%'
    BEGIN
	  IF @codmoneda = '01'
	   BEGIN
		 SET @condctacontable=' AND LTRIM(RTRIM(E.tdocumentocuentasoles)) LIKE ('''+@ctacontable+''') '
	   END
	  IF @codmoneda = '02'
	   BEGIN
		 SET @condctacontable= ' AND LTRIM(RTRIM(E.tdocumentocuentadolares)) LIKE ('''+@ctacontable+''') '
	   END
	END
END

declare @cadtmp varchar(2000)
set @cadtmp='SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
		B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
		B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
		B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla
  		INTO ##tmp_saldodoc' +@compu+ '  
	FROM 
		[' +@base+ '].dbo.vt_abono B,
		[' +@base+ '].dbo.cc_tipodocumento C
	where abonocancli =''*'''

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldodoc'+@compu)
  exec('DROP TABLE ##tmp_saldodoc'+@compu)

if @acuenta='1'
BEGIN
  exec(@cadtmp)
END

if @acuenta='0'
BEGIN
	Set @sqlcad=' 
 		SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
			B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,B.abonocanndqc,B.abonocanmoncan,
			B.abonocanimpcan,B.abonocanimpsol,B.abonocanfecan,
        	B.abonocanmoneda,B.abonocanimcan,B.abonocanforcan,
			B.abonocancuenta,B.abononumplanilla,B.abonotipoplanilla
  		INTO ##tmp_saldodoc' +@compu+ '  
		FROM 
			[' +@base+ '].dbo.vt_abono B,
			[' +@base+ '].dbo.cc_tipodocumento C,
 				(select clientecodigo,documentocargo,cargonumdoc from [' +@base+ '].dbo.vt_cargo D, [' +@base+ '].dbo.cc_tipodocumento E
 						where cargoapefecemi<=''' +@fecha+  '''  AND  cargoapeflgcan=0 AND
							cargoapeflgreg IS NULL ' +@condctacontable+ ' AND
							D.documentocargo=E.tdocumentocodigo ) AS Z
 		WHERE 	B.abonocancli=Z.clientecodigo AND
        		B.documentoabono=Z.documentocargo AND 
				B.abononumdoc=Z.cargonumdoc AND
				B.abonocanfecan<='''  +@fecha+ ''' AND 
				B.abonocantdqc=C.tdocumentocodigo AND
				B.abonocancli like ''' +@codcliente+   ''''
  	exec(@sqlcad)

END

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_saldoactualizado'+@compu)
exec('DROP TABLE ##tmp_saldoactualizado' +@compu )

set @sqlcad='
	SELECT 	A.clientecodigo, A.documentocargo, A.cargonumdoc,A.cargoapefecemi,
			A.cargoapefecvct,A.bancocodigo,A.monedacodigo,
			cargoapeimpape=ISNULL( dbo.tipodoc(E.tdocumentotipo,A.cargoapeimpape) ,0 ),
			cargoapeimppag=isnull(A.cargoapeimppag,0),
			cargopagadoux=isnull(A.cargoapeimppag,0),
			A.cargoapeflgcan,A.cargoapecarabo,
 			Y.*,
 			E.tdocumentodescripcion,G.bancodescripcion,
			I.clienteruc,I.clienterazonsocial,
			H.monedasimbolo
    INTO 	##tmp_saldoactualizado' +@compu+ '
	FROM 	[' +@base+ '].dbo.vt_cargo A, ##tmp_saldodoc' +@compu+ ' Y,
 			[' +@base+ '].dbo.cc_tipodocumento E,
			[' +@base+ '].dbo.gr_banco G,
			[' +@base+ '].dbo.gr_moneda H,
			[' +@base+ '].dbo.vt_cliente I
	WHERE 	A.clientecodigo*=Y.abonocancli AND
        	A.documentocargo*=Y.documentoabono AND 
			A.cargonumdoc*=Y.abononumdoc AND
	    	A.documentocargo=E.tdocumentocodigo AND
			A.bancocodigo*=G.bancocodigo AND
			A.monedacodigo=H.monedacodigo AND
        	A.clientecodigo=I.clientecodigo AND
			A.clientecodigo like ''' +@codcliente+ ''' AND 
			A.monedacodigo like ''' +@codmoneda+ '''	AND
            A.cargoapefecemi<=''' +@fecha+ ''' AND
			A.cargoapeflgcan=0  AND
			A.cargoapeflgreg IS NULL ' +@condctacontable + '
	ORDER BY cast(A.clientecodigo as int), A.documentocargo,A.cargonumdoc'

exec(@sqlcad)


IF @acuenta='0' 
BEGIN
	set @sqlcad='
		UPDATE ##tmp_saldoactualizado' +@compu+ ' SET ##tmp_saldoactualizado' +@compu+ '.cargopagadoux=B.saldoactual
			FROM ##tmp_saldoactualizado' +@compu+ ' A,
				 (SELECT clientecodigo,documentocargo,cargonumdoc,saldoactual=SUM(ISNULL(abonocanimpsol,0)) 
					FROM ##tmp_saldoactualizado' +@compu+ ' 
 					GROUP BY clientecodigo,documentocargo,cargonumdoc ) as B
			WHERE A.clientecodigo=B.clientecodigo AND A.documentocargo=B.documentocargo AND
	  			A.cargonumdoc=B.cargonumdoc'

	exec(@sqlcad)
END

IF @fecha<convert(varchar(10),getdate(),103) AND @acuenta='1'
BEGIN
	exec('UPDATE ##tmp_saldoactualizado' +@compu+ ' SET cargopagadoux=0')
		
	SET @sqlcad=''
	SET @sqlcad='
	UPDATE ##tmp_saldoactualizado' +@compu+ ' SET ##tmp_saldoactualizado' +@compu+ '.cargopagadoux=Y.saldoactual
				FROM ##tmp_saldoactualizado' +@compu+ ' A,
			(SELECT B.abonocancli,B.documentoabono,B.abononumdoc,saldoactual=SUM(ISNULL(B.abonocanimpsol,0))
			FROM 
				[' +@base+ '].dbo.vt_abono B,
				[' +@base+ '].dbo.cc_tipodocumento C,
 					(select clientecodigo,documentocargo,cargonumdoc from [' +@base+ '].dbo.vt_cargo D, [' +@base+ '].dbo.cc_tipodocumento E
	 							where cargoapefecemi<=''' +@fecha+  '''  AND  cargoapeflgcan=0 AND
								cargoapeflgreg IS NULL ' +@condctacontable+ ' AND
								D.documentocargo=E.tdocumentocodigo ) AS Z
	 		WHERE 	B.abonocancli=Z.clientecodigo AND
    	    		B.documentoabono=Z.documentocargo AND 
					B.abononumdoc=Z.cargonumdoc AND
					B.abonocanfecan<='''  +@fecha+ ''' AND 
					B.abonocantdqc=C.tdocumentocodigo AND
					B.abonocancli like ''' +@codcliente+   '''
        	GROUP BY B.abonocancli,B.documentoabono,B.abononumdoc) as Y
		WHERE A.clientecodigo=Y.abonocancli AND A.documentocargo=Y.documentoabono AND
	  		A.cargonumdoc=Y.abononumdoc'
	exec(@sqlcad)

END

set @sqlcad=''
set @sqlcad='SELECT cod_documento=a.documentocargo,a.monedacodigo,desc_documento=b.tdocumentodescripcion,
	SALDO_SOLES = CASE 
	WHEN a.monedacodigo = 01 THEN SUM(isnull(a.cargoapeimpape,0)) - SUM(isnull(a.cargopagadoux,0))
    ELSE 0
	end,
	SALDO_DOLARES = CASE 
	WHEN a.monedacodigo = 02 THEN SUM(isnull(a.cargoapeimpape,0)) - SUM(isnull(a.cargopagadoux,0))
    ELSE 0
	end
FROM  ##tmp_saldoactualizado' +@compu+  ' a, [' +@base+ '].dbo.cc_tipodocumento b
WHERE a.documentocargo=b.tdocumentocodigo
GROUP BY documentocargo,monedacodigo,b.tdocumentodescripcion'

exec(@sqlcad)

set nocount off

--select * from ##tmp_saldodocdesarrollo3 order by abonocancli,documentoabono,abononumdoc
--exec cc_EMLB_SubSaldoxCliente_Detalle 'ventas_prueba','DESARROLLO3','18/11/2002','1','%','%','%'








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

