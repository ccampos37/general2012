SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO
















ALTER           proc cc_EMLB_CtaCtexCliente(
@base varchar(50),
@compu varchar(20),
@fechaini varchar(10),
@fechafin varchar(10),
@fecha varchar(10),
@codmoneda varchar(2),
@codcliente varchar(50),
@codzona varchar(10)
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
	WHERE abonocancli =''*'''

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
		A.zonacodigo,
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
				A.zonacodigo LIKE ''' +@codzona+ ''' AND
	        	floor(cast(A.cargoapefecemi as real)) <=' + cast(dbo.fn_datenumber(day(@fecha),month(@fecha),year(@fecha)) as varchar(20))  + ' AND
				A.cargoapeflgcan=0  AND
				A.cargoapeflgreg IS NULL
	ORDER BY A.clientecodigo, A.documentocargo,A.cargonumdoc'

exec(@sqlcad)

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
					cargoapeflgreg IS NULL AND	D.documentocargo=E.tdocumentocodigo ) AS Z
	 		WHERE 	B.abonocancli=Z.clientecodigo AND
	    	    		B.documentoabono=Z.documentocargo AND 
						B.abononumdoc=Z.cargonumdoc AND
						FLOOR(CAST (B.abonocanfecan AS REAL)) <='  + CAST(DBO.fn_datenumber(DAY(@fecha),MONTH(@FECHA),YEAR(@FECHA)) AS VARCHAR(20)) + ' AND 
						B.abonocantdqc=C.tdocumentocodigo AND
						B.abonocancli like ''' +@codcliente+ '''
        	GROUP BY B.abonocancli,B.documentoabono,B.abononumdoc) as Y
		WHERE	A.clientecodigo=Y.abonocancli AND A.documentocargo=Y.documentoabono AND
	  		A.cargonumdoc=Y.abononumdoc'
	exec(@sqlcad)

exec('UPDATE ##tmp_saldoinicial' +@compu+ ' SET saldoinicial=(cargoapeimpape-cargopagadoux)')

IF EXISTS (SELECT NAME FROM tempdb.dbo.sysobjects WHERE NAME='##tmp_abonos'+@compu)
exec('DROP TABLE ##tmp_abonos' +@compu )

declare @cadsql nvarchar(4000)
SET @cadsql='
	SELECT B.abonocancli,B.documentoabono,B.abononumdoc,B.abonocancarabo,
			B.abonocantdqc,tipodescripcion=C.tdocumentodescripcion,
			B.abonocanndqc,B.abonocanmoncan,B.abonocanimpcan,
			B.abonocanimpsol,B.abonocanfecan,B.abonocanmoneda,
			B.abonocanimcan,B.abonocanforcan,B.abonocancuenta,
			B.abononumplanilla,B.abonotipoplanilla,
			simbmonabo=D.monedasimbolo,
			B.abonocantipcam,
			B.abonocanbco

		INTO ##tmp_abonos' +@compu+ '
		FROM 
			[' +@base+ '].dbo.vt_abono B,
			[' +@base+ '].dbo.cc_tipodocumento C,
			[' +@base+ '].dbo.gr_moneda D
 		WHERE 	
			B.abonocanfecan between ''' +@fechaini+ ''' AND ''' +@fechafin+ ''' AND
			B.abonocantdqc=C.tdocumentocodigo AND
			B.abonocanmoncan=D.monedacodigo AND 
			B.abonocancli LIKE ''' +@codcliente+   ''' AND
			B.zonacodigo LIKE ''' +@codzona+ ''''
exec(@cadsql)

SET @cadsql='
	SELECT 	AA.*,
			C.clienteruc,C.clienterazonsocial,
			descdoccargo=D.tdocumentodescripcion,
			numplanillacargo=E.abononumplanilla,
			banco=F.bancodescripcion
		FROM	
			[' +@base+ '].dbo.vt_cliente C,
			[' +@base+ '].dbo.cc_tipodocumento D,
			[' +@base+ '].dbo.vt_cargo E,
			[' +@base+ '].dbo.gr_banco F,

			(SELECT A.clientecodigo, A.documentocargo, A.cargonumdoc, B.abononumplanilla, A.cargoapefecemi,A.cargoapefecvct, A.monedacodigo, A.monedasimbolo, A.cargoapeimpape,
				   	B.abonocantdqc,B.tipodescripcion,B.abonocanndqc,B.abonocanmoncan,abonocanimpcan=ISNULL(B.abonocanimpcan,0),abonocanimpsol=ISNULL(B.abonocanimpsol,0),
			   		B.abonocanfecan, B.simbmonabo,
			  			SaldoInicial,
						A.zonacodigo,
						B.abonocanbco
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
						SaldoInicial=cast(0 as numeric(25,9)),
						A.zonacodigo,
						B.abonocanbco
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
			AA.abonocanbco*=F.bancocodigo AND
			AA.clientecodigo LIKE ''' +@codcliente+ ''' AND
			AA.monedacodigo  LIKE ''' +@codmoneda+ ''''
			
exec(@cadsql)

set nocount off

--select * from ##tmp_saldodocdesarrollo3 order by abonocancli,documentoabono,abononumdoc
--exec cc_EMLB_CtaCtexCliente 'ventas_prueba','DESARROLLO3','01/12/2002','28/12/2002','30/11/2002','%','%','%'

--select * from ventas_prueba.dbo.vt_zona
--zonacodigo zonadescripcion
--select * from ventas_prueba.dbo.vt_abono





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

