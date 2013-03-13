/*
		timporte=
			case when ZZ.monedaapertura='02' then
				ZZ.ImporteApertura*isnull((select tipocambioventa from [Contaprueba].dbo.ct_tipocambio as M where M.tipocambiofecha=ZZ.FechaEmision),1)
			else
				ZZ.ImporteApertura
			end
*/


Declare @CadCtasEgreso varchar(5000)
Declare @Base varchar(50)
Declare @BaseConta varchar(50)
Declare @NombrePC varchar(50)


--select * from te_cabecerarecibos
--select * from te_conceptocaja

set @Base='Ventas_Prueba'
set @BaseConta='Contaprueba'
set @NombrePC='Desarrollo3'

if exists (select name from tempdb.dbo.sysobjects where name='##tmpEgreso' +@NombrePC) 
  exec('DROP TABLE ##tmpEgreso' +@NombrePC)

if exists (select name from tempdb.dbo.sysobjects where name='##tmpCuentaEgreso' + @NombrePC)
  exec('DROP TABLE ##tmpCuentaEgreso' +@NombrePC)

set @CadCtasEgreso='

select	ZZ.*,
			tcemision=
				case when ZZ.monedaapertura=''02''
					then	isnull((select tipocambioventa from '  +@BaseConta+ '.dbo.ct_tipocambio as M where M.tipocambiofecha=ZZ.FechaEmision),1)
					else	1
				end
into ##tmpEgreso' +@NombrePC+ '
from 
(select 	
	A.cabrec_numrecibo,A.detrec_item,
	ImporteSoles=A.detrec_importesoles,
	ImporteDolar=A.detrec_importedolares,A.detrec_monedacancela,
	A.detrec_cajabanco1,A.detrec_numctacte,A.detrec_tipocajabanco,
	concepto=A.detrec_tipodoc_concepto,A.detrec_tdqc,A.detrec_ndqc,
	A.detrec_observacion,A.detrec_fechacancela,tccancela=3.550,
	A.detrec_numdocumento,			
	A.detrec_tipodoc_concepto,
	C.operacioncontrolaclienteprov,
	B.clientecodigo,
	FechaEmision=
  		isnull(
			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),''X'')) 
     			When ''P'' then (select cp.cargoapefecemi from [ventas_prueba].dbo.cp_cargo cp where cp.documentocargo=A.detrec_tipodoc_concepto and 
									cp.cargonumdoc=A.detrec_numdocumento and cp.clientecodigo like B.clientecodigo)
				When ''C'' then (select cc.cargoapefecemi from [ventas_prueba].dbo.vt_cargo cc where cc.documentocargo=A.detrec_tipodoc_concepto and 
									cc.cargonumdoc=A.detrec_numdocumento and cc.clientecodigo like B.clientecodigo)
  				Else
					B.cabrec_fechadocumento
			End,B.cabrec_fechadocumento),
/*
	ImporteApertura=
  		isnull(
			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),''X'')) 
     			When ''P'' then (select cp.cargoapeimpape from [ventas_prueba].dbo.cp_cargo cp where cp.documentocargo=A.detrec_tipodoc_concepto and 
									cp.cargonumdoc=A.detrec_numdocumento and cp.clientecodigo like B.clientecodigo)
				When ''C'' then (select cc.cargoapeimpape from [ventas_prueba].dbo.vt_cargo cc where cc.documentocargo=A.detrec_tipodoc_concepto and 
									cc.cargonumdoc=A.detrec_numdocumento and cc.clientecodigo like B.clientecodigo)
  				End,0),
*/
	MonedaApertura=
  		isnull(
			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),''X'')) 
     			When ''P'' then (select cp.monedacodigo from [ventas_prueba].dbo.cp_cargo cp where cp.documentocargo=A.detrec_tipodoc_concepto and 
									cp.cargonumdoc=A.detrec_numdocumento and cp.clientecodigo like B.clientecodigo)
				When ''C'' then (select cc.monedacodigo from [ventas_prueba].dbo.vt_cargo cc where cc.documentocargo=A.detrec_tipodoc_concepto and 
									cc.cargonumdoc=A.detrec_numdocumento and cc.clientecodigo like B.clientecodigo)
  				End,''00''),

	cuenta=
		case when detrec_monedadocumento=''01'' then
	     	Isnull(
   	  			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),''X'')) 
      		    	When ''P'' then (Select P.tdocumentocuentasoles  from  [ventas_prueba].dbo.cp_tipodocumento P Where P.tdocumentocodigo=a.detrec_tipodoc_concepto)
        	  			When ''C'' then (Select Cl.tdocumentocuentasoles from  [ventas_prueba].dbo.cc_tipodocumento Cl Where Cl.tdocumentocodigo=a.detrec_tipodoc_concepto)           
        			Else  (select conceptocuentasoles from [ventas_prueba].dbo.te_conceptocaja where conceptocodigo=a.detrec_tipodoc_concepto)
       			End,''XXXXXXXXXXX'') 
		Else
	     	Isnull(
   	  			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),''X'')) 
      		    	When ''P'' then (Select P.tdocumentocuentadolares  from [ventas_prueba].dbo.cp_tipodocumento P Where P.tdocumentocodigo=a.detrec_tipodoc_concepto)
        	  			When ''C'' then (Select Cl.tdocumentocuentadolares  from [ventas_prueba].dbo.cc_tipodocumento Cl Where Cl.tdocumentocodigo=a.detrec_tipodoc_concepto)           
        			Else  (select conceptocuentadolar from [ventas_prueba].dbo.te_conceptocaja where conceptocodigo=a.detrec_tipodoc_concepto)
       			End,''YYYYYYYYYYY'') 
		end
from 
	[' +@Base+ '].dbo.te_detallerecibos A,
	[' +@Base+ '].dbo.te_cabecerarecibos B,
	[' +@Base+ '].dbo.te_operaciongeneral C
where 
	a.cabrec_numrecibo=b.cabrec_numrecibo and
	b.operacioncodigo*=c.operacioncodigo and 
	b.cabrec_ingsal like ''%'') as ZZ '


exec(@CadCtasEgreso)


set @CadCtasEgreso=''
set @CadCtasEgreso='
select
	 	cabrec_numrecibo,detrec_item,MonedaApertura as MonedaCodigo,
		detrec_tipodoc_concepto,detrec_numdocumento,''000000000'' as clientecodigo,
		cuenta,detrec_observacion,
		detrec_fechacancela as FechaCancela,
		tccancela as TipoCambio,
		DebeS=
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then 
					case when detrec_monedacancela=''02'' 
						then	ImporteDolar*tcemision
						else	ImporteSoles
					end
  	  			When ''C'' then 0.00           
  			Else  0.00
  			End,
		HaberS=
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then 0.00
  	  			When ''C'' then 
					case when detrec_monedacancela=''02'' 
						then	ImporteDolar*tcemision
						else	ImporteSoles
					end
  			Else  0.00
  			End,

		DebeD=
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then ImporteDolar
  	  			When ''C'' then 0.00           
  			Else  0.00
  			End,
		HaberD=
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then 0.00
  	  			When ''C'' then ImporteDolar           
  			Else  0.00
  			End,

		ImporteSol=
			case when detrec_monedacancela=''02'' 
				then	ImporteDolar*tcemision
				else	ImporteSoles
			end,
		ImporteDolar,
      ImporteSoles
into ##tmpCuentaEgreso'+ @NombrePC+ '
from 
	##tmpEgreso'+ @NombrePC+ '
order by 1'



--select * from ##tmpEgresoDesarrollo3
/*
cabrec_numrecibo detrec_item MonedaCodigo detrec_tipodoc_concepto detrec_numdocumento clientecodigo cuenta               detrec_observacion                                 FechaCancela                                           TipoCambio DebeS                                                 HaberS                                                DebeD                                                 HaberD                                                ImporteSol                                            ImporteDolar                                          ImporteSoles                                          
---------------- ----------- ------------ ----------------------- ------------------- ------------- -------------------- -------------------------------------------------- ------------------------------------------------------ ---------- ----------------------------------------------------- ----------------------------------------------------- ----------------------------------------------------- ----------------------------------------------------- ----------------------------------------------------- ----------------------------------------------------- ----------------------------------------------------- 
100001           1           00           80                      546                 000000000     123480               TRANSFERENCIA                                      2002-12-04 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   20000.0                                               20000.0                                               69600.0
100002           1           00           80                      04422249            000000000     123480               PAGO DE LUZ Y AGUA                                 2002-12-03 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   82.5                                                  82.5                                                  82.5
100003           1           00           80                      353                 000000000     123480               TRANSFERENCIA                                      2002-12-04 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   15000.0                                               15000.0                                               52500.0
100004           1           00           80                      459                 000000000     123480               TRANSFERENCIA                                      2002-12-10 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   10000.0                                               10000.0                                               35000.0
100005           1           00           80                      4422258             000000000     123480               PAGO AFP NOVIEMBRE                                 2002-12-06 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   14090.48                                              14090.48                                              14090.48
100006           1           00           80                      7490092             000000000     123480               PAGO CTS NOV.                                      2002-12-06 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   975.32000000000005                                    975.32000000000005                                    3413.6200000000003
100007           1           00           80                      0                   000000000     123480                                                                  2003-02-21 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   285.71428571428572                                    285.71428571428572                                    1000.0
100008           1           00           80                      2365                000000000     123480               De Soles a Dolares                                 2003-02-21 00:00:00.000                                3.550      0.0                                                   0.0                                                   0.0                                                   0.0                                                   3428.5714285714284                                    3428.5714285714284                                    12000.0
200001           1           01           01                      0011466             000000000     421101               SERVICIO DE ESTAMPADO                              2003-01-15 00:00:00.000                                3.550      2940.3000000000002                                    0.0                                                   843.21766561514198                                    0.0                                                   2940.3000000000002                                    843.21766561514198                                    2940.3000000000002
200002           1           02           60                      94600117429         000000000     423101               CANCELAC. DE L/. 9460117119 DE CLARIANT SUIZA      2002-12-10 00:00:00.000                                3.550      18958.720000000001                                    0.0                                                   5386.0                                                0.0                                                   19400.371999999999                                    5386.0                                                18958.720000000001
200003           2           02           01                      12100090753         000000000     421101               SILVATOL S.A                                       2002-12-13 00:00:00.000                                3.550      1489.3488                                             0.0                                                   424.80000000000001                                    0.0                                                   1529.7048                                             424.80000000000001                                    1489.3488
200004           1           02           01                      20100032713         000000000     421101               CANC. POR UKOSOFT 49                               2002-12-13 00:00:00.000                                3.550      1064.5999999999999                                    0.0                                                   306.80115273775209                                    0.0                                                   1064.5999999999999                                    306.80115273775209                                    1064.5999999999999
200005           1           02           01                      20100032767         000000000     421101               CANC. POR UKOSOFT 49                               2002-12-13 00:00:00.000                                3.550      2129.1999999999998                                    0.0                                                   613.60230547550418                                    0.0                                                   2129.1999999999998                                    613.60230547550418                                    2129.1999999999998
200006           1           02           01                      20100032678         000000000     421101               CANC POR UKOSOFT 49                                2002-12-13 00:00:00.000                                3.550      533.83000000000004                                    0.0                                                   153.39942528735634                                    0.0                                                   533.83000000000004                                    153.39942528735634                                    533.83000000000004
200007           1           01           01                      00100000075         000000000     421101               SERV.ALQUILER DE MONTACARGAS                       2002-12-13 00:00:00.000                                3.550      64.900000000000006                                    0.0                                                   18.511123787792357                                    0.0                                                   64.900000000000006                                    18.511123787792357                                    64.900000000000006
200007           2           01           01                      00100000076         000000000     421101               SERV.ALQUILER DE MONTACARGAS                       2002-12-13 00:00:00.000                                3.550      64.900000000000006                                    0.0                                                   18.511123787792357                                    0.0                                                   64.900000000000006                                    18.511123787792357                                    64.900000000000006
200007           3           01           01                      00100000078         000000000     421101               SERV.ALQUILER DE MONTACARGAS                       2002-12-13 00:00:00.000                                3.550      129.80000000000001                                    0.0                                                   37.022247575584714                                    0.0                                                   129.80000000000001                                    37.022247575584714                                    129.80000000000001
200008           1           01           01                      00100000440         000000000     421101               CANC. POR PARIHUELAS DE MADERA                     2002-12-13 00:00:00.000                                3.550      1793.5999999999999                                    0.0                                                   511.5801483171706                                     0.0                                                   1793.5999999999999                                    511.5801483171706                                     1793.5999999999999
200009           1           02           01                      00100080236         000000000     421101               POR CANC DE 62042RSRC3                             2002-12-13 00:00:00.000                                3.550      105.91625999999999                                    0.0                                                   30.210000000000001                                    0.0                                                   108.84663                                             30.210000000000001                                    105.91625999999999
200010           1           02           01                      00100032387         000000000     421101               CANC. DE FAJA OPTIBEL 7M 1180                      2002-12-13 00:00:00.000                                3.550      374.33561999999995                                    0.0                                                   106.77                                                0.0                                                   382.77044999999998                                    106.77                                                374.33561999999995
200011           1           02           01                      00100032449         000000000     421101               CANC. DE FAJA OPTIBEL 440 T 10                     2002-12-13 00:00:00.000                                3.550      686.75527999999997                                    0.0                                                   195.88                                                0.0                                                   701.25040000000001                                    195.88                                                686.75527999999997
200012           1           01           01                      00300006688         000000000     421101               CANC. DE UTILES DE OFICNA                          2002-12-13 00:00:00.000                                3.550      852.13                                                0.0                                                   243.04905875641759                                    0.0                                                   852.13                                                243.04905875641759                                    852.13
200012           2           01           01                      00400004853         000000000     421101               CANC. DE UTILES DE OFICNA                          2002-12-13 00:00:00.000                                3.550      127.34999999999999                                    0.0                                                   36.323445521962348                                    0.0                                                   127.34999999999999                                    36.323445521962348                                    127.34999999999999
200012           3           01           01                      00300006703         000000000     421101               CANC. DE UTILES DE OFICNA                          2002-12-13 00:00:00.000                                3.550      741.44000000000005                                    0.0                                                   211.4774671990873                                     0.0                                                   741.44000000000005                                    211.4774671990873                                     741.44000000000005
200013           1           02           01                      00100011819         000000000     421101               CANC. DE JERSEY 20/1 X2 C TEJIDOS                  2002-12-13 00:00:00.000                                3.550      646.75181999999995                                    0.0                                                   184.47                                                0.0                                                   665.93669999999997                                    184.47                                                646.75181999999995
200013           2           02           01                      00100011907         000000000     421101               CANC. DE JERSEY 24/1 X2 C20/1 X2 C TEJIDOS         2002-12-13 00:00:00.000                                3.550      7615.7682599999998                                    0.0                                                   2172.21                                               0.0                                                   7837.3336800000006                                    2172.21                                               7615.7682599999998
*/