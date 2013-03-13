UPDATE	vt_cargo set
			cargoapeimppag=cargoapeimppag-YY.abonocanimpsol,
			cargoapeflgcan=0 
--SELECT AA.* 
From	vt_cargo AA,
		(select A.documentoabono,A.abononumdoc,A.abonocancli,A.abonotipoplanilla,A.abonocanmoneda,
			 A.abonocanimcan,A.abonocanforcan,A.abonocanimpsol 
 				from 	ventas_prueba.dbo.vt_abono A,
						(select 	A.* from [ventas_prueba].dbo.vt_cargo A,
										[ventas_prueba].dbo.cc_tipoplanilla d
							where  
								--isnull(cargoapeflgreg,0)<>1 and
								A.abonotipoplanilla=d.tplanillacodigo and  d.tplanillacanjes='1' and
								A.clientecodigo like '65' and
								A.documentocargo='60' and 
								A.cargonumdoc like '00000001007') as ZZ
		where A.abononumplanilla=ZZ.abononumplanilla and 
				A.abonotipoplanilla=ZZ.abonotipoplanilla and
				A.abononumplanilla like '%' ) as YY
where AA.documentocargo=YY.documentoabono and
		AA.cargonumdoc=YY.abononumdoc	and
		AA.clientecodigo=YY.abonocancli and 
		AA.cargoapeimppag<>0 

