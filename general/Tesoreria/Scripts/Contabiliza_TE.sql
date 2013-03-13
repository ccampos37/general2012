select 
		b.cabrec_numrecibo,
		b.clientecodigo,
		a.detrec_tipodoc_concepto,
		a.detrec_numdocumento,
		g.cargoapefecemi as FecEmision,
		g.monedacodigo as MonedaOrigen,
		g.cargoapeimpape as CargoOrigen,
		tcemision=
			case when g.monedacodigo='02' then
	    		isnull((select tipocambioventa from [Contaprueba].dbo.ct_tipocambio as M where M.tipocambiofecha=g.cargoapefecemi),0)  
			else
				1
			end,	
/*
		tccancela=
			case when e.abonocanmoncan=''02'' then
				isnull((select tipocambiocompra from [' +@basecontab+ '].dbo.ct_tipocambio as M where M.tipocambiofecha=e.abonocanfecan ),0)
			else
				1
			end,
*/
		cuenta42=case when g.monedacodigo='01' then h.tdocumentocuentasoles else h.tdocumentocuentadolares end,

		timporte=g.cargoapeimpape*tcemision,
		a.detrec_fechacancela,
  	   a.detrec_emisioncheque,
		detrec_monedacancela, 
		detrec_tipocajabanco,a.detrec_cajabanco1, a.detrec_numctacte,
		b.cabrec_ingsal,
		a.detrec_fechacancela,
		a.detrec_cajabanco1+a.detrec_numctacte as Codigo,
		a.detrec_monedadocumento,
		a.detrec_numdocumento,
		DescCajaBanco= case when a.detrec_tipocajabanco='B' then d.bancodescripcion else e.cajadescripcion end,
		a.detrec_forcan,
		a.detrec_tdqc,
		a.detrec_ndqc,
		0 as SaldoInicial,
      f.monedasimbolo,
  		ProveCliConc=
       	Isnull(
     			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),'X')) 
          	When 'P' then (Select Top 1 P.clienterazonsocial  from [ventas_prueba].dbo.cp_proveedor P Where P.clientecodigo=b.clientecodigo)
        	  	When 'C' then (Select Top 1 Cl.clienterazonsocial  from  [ventas_prueba].dbo.vt_cliente Cl Where Cl.clientecodigo=b.clientecodigo)           
        		Else  b.cabrec_descripcion
       		End,'') ,       
  		Td_Concep=
	      Isnull(
		   	case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),'X')) 
	       	when 'P' then (Select X.tdocumentodescripcion from  [ventas_prueba].dbo.cp_tipodocumento X Where X.tdocumentocodigo=A.detrec_tipodoc_concepto)
	        	When 'C' then (Select Y.tdocumentodescripcion from  [ventas_prueba].dbo.cc_tipodocumento Y Where Y.tdocumentocodigo=A.detrec_tipodoc_concepto)           
        	  	Else  (Select G.conceptodescripcion  from [ventas_prueba].dbo.te_conceptocaja G  where G.conceptocodigo=A.detrec_tipodoc_concepto)
       		End,'')
		from  [ventas_prueba].dbo.te_detallerecibos a, 
				[ventas_prueba].dbo.te_cabecerarecibos b, 
				[ventas_prueba].dbo.te_operaciongeneral c,
				[ventas_prueba].dbo.gr_banco d,
				[ventas_prueba].dbo.te_codigocaja e,
				[ventas_prueba].dbo.gr_moneda f,
				[ventas_prueba].dbo.cp_cargo g,
				[ventas_prueba].dbo.cp_tipodocumento h 
				

		where a.cabrec_numrecibo=b.cabrec_numrecibo and 
				b.operacioncodigo=c.operacioncodigo and 
				ltrim(rtrim(a.detrec_cajabanco1+a.detrec_numctacte)) like  '%'  and
				detrec_tipocajabanco like 'B' and 
				a.detrec_cajabanco1*=d.bancocodigo and
				a.detrec_cajabanco1*=e.cajacodigo  and 
				a.detrec_monedacancela=f.monedacodigo and

				b.clientecodigo*=g.clientecodigo and
				a.detrec_tipodoc_concepto*=g.documentocargo and
				a.detrec_numdocumento*=g.cargonumdoc	and
				a.detrec_tipodoc_concepto*= h.tdocumentocodigo


--select * from cp_cargo


				b.cabrec_ingsal='E' and
				month(a.detrec_fechacancela)=12 and year(a.detrec_fechacancela)=2002
		order by a.detrec_cajabanco1, a.detrec_numctacte,a.detrec_fechacancela
