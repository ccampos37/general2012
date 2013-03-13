
Alter Procedure te_GeneraTempContable(

@Base varchar(50),
@BaseConta varchar(50),
@CtaGanCambio varchar(20),
@CtaPerCambio varchar(20),
@NombrePC varchar(50),
@MesProceso varchar(2),
@AnnoProceso varchar(4),
@TipoMov varchar(1) ) /*I:Ingreso E:Egreso */
as


/*
set @Base='Ventas_Prueba'
set @BaseConta='Contaprueba'
set @NombrePC='Desarrollo3'
set @CtaGanCambio='776101'
set @CtaPerCambio='976101'
*/


Declare @CadCtasEgreso varchar(5000)
Declare @CadCtasCaja varchar(5000)
Declare @CadDifCambio varchar(4000)

set nocount on

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
	ImporteApertura=
  		isnull(
			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),''X'')) 
     			When ''P'' then (select cp.cargoapeimpape from [ventas_prueba].dbo.cp_cargo cp where cp.documentocargo=A.detrec_tipodoc_concepto and 
									cp.cargonumdoc=A.detrec_numdocumento and cp.clientecodigo like B.clientecodigo)
				When ''C'' then (select cc.cargoapeimpape from [ventas_prueba].dbo.vt_cargo cc where cc.documentocargo=A.detrec_tipodoc_concepto and 
									cc.cargonumdoc=A.detrec_numdocumento and cc.clientecodigo like B.clientecodigo)
  				End,0),
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
       			End,''423101'') 
		Else
	     	Isnull(
   	  			case upper(isnull(rtrim(ltrim(C.operacioncontrolaclienteprov)),''X'')) 
      		    	When ''P'' then (Select P.tdocumentocuentadolares  from [ventas_prueba].dbo.cp_tipodocumento P Where P.tdocumentocodigo=a.detrec_tipodoc_concepto)
        	  			When ''C'' then (Select Cl.tdocumentocuentadolares  from [ventas_prueba].dbo.cc_tipodocumento Cl Where Cl.tdocumentocodigo=a.detrec_tipodoc_concepto)           
        			Else  (select conceptocuentadolar from [ventas_prueba].dbo.te_conceptocaja where conceptocodigo=a.detrec_tipodoc_concepto)
       			End,''423101'') 
		end
from 
	[' +@Base+ '].dbo.te_detallerecibos A,
	[' +@Base+ '].dbo.te_cabecerarecibos B,
	[' +@Base+ '].dbo.te_operaciongeneral C
where 
	a.cabrec_numrecibo=b.cabrec_numrecibo and
	b.operacioncodigo*=c.operacioncodigo and
	month(a.detrec_fechacancela)=''' +@MesProceso+ ''' and
	year(a.detrec_fechacancela)=''' +@AnnoProceso+ ''' and
	b.cabrec_ingsal like ''' +@TipoMov+''') as ZZ '

exec(@CadCtasEgreso)


set @CadCtasEgreso=''
set @CadCtasEgreso='
select
	 	cabrec_numrecibo,detrec_item,MonedaApertura as MonedaCodigo,
		detrec_tipodoc_concepto,detrec_numdocumento,
		clientecodigo,
		cuenta,detrec_observacion,
		FechaEmision,
		detrec_fechacancela as FechaCancela,
		tccancela as TipoCambio,
		DebeS=Cast(Round(
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then 
					case when detrec_monedacancela=''02'' 
						then	ImporteDolar*tcemision
						else	ImporteSoles
					end
  	  			When ''C'' then 0.00           
  			Else  0.00
  			End,2) as Numeric(15,2)),
		HaberS=Cast(Round(
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then 0.00
  	  			When ''C'' then 
					case when detrec_monedacancela=''02'' 
						then	ImporteDolar*tcemision
						else	ImporteSoles
					end
  			Else  0.00
  			End,2) as numeric(15,2)),

		DebeD=Cast( Round(
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then ImporteDolar
  	  			When ''C'' then 0.00           
  			Else  0.00
  			End,2) as Numeric(15,2)),
		HaberD=Cast( Round(
  			case upper(isnull(rtrim(ltrim(operacioncontrolaclienteprov)),''X'')) 
  		    	When ''P'' then 0.00
  	  			When ''C'' then ImporteDolar           
  			Else  0.00
  			End,2) as Numeric(15,2)),

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

Exec(@CadCtasEgreso)



if exists (Select name from tempdb.dbo.sysobjects where name='##tmpCuentaCaja' +@NombrePC)
  exec('drop table ##tmpCuentaCaja' +@NombrePC)

set @CadCtasCaja='
select
	 	cabrec_numrecibo,
		detrec_item=
		(select max(detrec_item) from ##tmpEgreso' +@NombrePC+ ' A where cabrec_numrecibo=TT.cabrec_numrecibo)+1,
		detrec_monedacancela as MonedaCodigo,
		detrec_tipodoc_concepto,
		detrec_numdocumento,
		clientecodigo,
		cuenta,detrec_observacion,
		detrec_fechacancela as FechaEmision,
		detrec_fechacancela as FechaCancela,tccancela as TipoCambio,
		DebeS=0.00,
		HaberS=
			case when detrec_monedacancela=''02'' 
				then	Cast(Round(ImporteDolar*tccancela,2) as Numeric(15,2))
				else	Cast(Round(ImporteSoles,2) as Numeric(15,2))
			end,
		DebeD=0.00,
		HaberD=Cast(Round(ImporteDolar,2) as Numeric(15,2)) ,
		ImporteSol=
			case when detrec_monedacancela=''02'' 
				then	Round(ImporteDolar*tccancela,2)
				else	Round(ImporteSoles,2)
			end,
		ImporteDolar,
      ImporteSoles

into ##tmpCuentaCaja' +@NombrePC+ '

from
(Select XX.*,  
		tccancela=
			case when XX.detrec_monedacancela=''02'' 
					then	isnull((select tipocambioventa from [' +@BaseConta+'].dbo.ct_tipocambio as M where M.tipocambiofecha=XX.detrec_fechacancela),0)
					else	1
			end
from
(select 	distinct ZZ.cabrec_numrecibo,
			ZZ.ImporteSoles,
			ZZ.ImporteDolar,
			YY.detrec_monedacancela, 
			YY.detrec_cajabanco1,YY.detrec_numctacte,
			YY.detrec_tipocajabanco,
			cuenta=
				case when YY.detrec_tipocajabanco=''C'' 
					then 
							case when detrec_monedacancela=''01'' 
								then
									(select C.cajacuentasoles from [' +@Base+ '].dbo.te_codigocaja C where C.cajacodigo=YY.detrec_cajabanco1)
								else
									(select C.cajacuentadolares from [' +@Base+ '].dbo.te_codigocaja C where C.cajacodigo=YY.detrec_cajabanco1)
							end
					else
						(select cbanco_cuenta from [' +@Base+ '].dbo.te_cuentabancos
							where cbanco_codigo=YY.detrec_cajabanco1 and monedacodigo=YY.detrec_monedacancela and cbanco_numero=YY.detrec_numctacte)
				end,
			detrec_tdqc=(select top 1 A.detrec_tdqc from [' +@Base+ '].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_ndqc=(select top 1 A.detrec_ndqc from [' +@Base+ '].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_observacion=(select top 1 A.detrec_observacion from [' +@Base+ '].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_fechacancela=(select top 1 A.detrec_fechacancela from [' +@Base+ '].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_tipodoc_concepto=(select top 1 detrec_tipodoc_concepto from [' +@Base+ '].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_numdocumento=(select top 1 detrec_numdocumento from [' +@Base+ '].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			clientecodigo=(select top 1 clientecodigo from [' +@Base+ '].dbo.te_cabecerarecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo)
from
	(select 	bb.cabrec_numrecibo,
			ImporteSoles=sum(bb.detrec_importesoles),
			ImporteDolar=sum(bb.detrec_importedolares) 
		from [' +@Base+ '].dbo.te_detallerecibos bb
		where month(bb.detrec_fechacancela)=''' +@mesproceso+ ''' and year(bb.detrec_fechacancela)=''' +@annoproceso+''' and detrec_tipodoc_concepto<>''80''
		group by bb.cabrec_numrecibo) as ZZ,
	[' +@Base+ '].dbo.te_detallerecibos YY	
where 
	ZZ.cabrec_numrecibo=YY.cabrec_numrecibo) as XX ) as TT
order by 1'

exec(@CadCtasCaja)


if exists (Select name from tempdb.dbo.sysobjects where name='##tmpAsientosConta' +@NombrePC)
  exec('drop table ##tmpAsientosConta' +@NombrePC)

set @CadDifCambio='
select * into ##tmpAsientosConta' +@NombrePC+ ' from ##tmpCuentaEgreso' +@NombrePC+ ' union all
	select * from ##tmpCuentaCaja' +@NombrePC+ ' union all
		select * from 
(select YY.cabrec_numrecibo,
detrec_item=(select max(detrec_item) from ##tmpCuentaCaja' +@NombrePC+ ' A where cabrec_numrecibo=YY.cabrec_numrecibo)+1  ,
HH.MonedaCodigo,HH.detrec_tipodoc_concepto, 
HH.detrec_numdocumento,HH.clientecodigo,   
cuenta= 
  case when YY.diferencia>0 then ''' +@CtaGanCambio+  ''' else ''' +@CtaPerCambio+ ''' end,
HH.detrec_observacion,
HH.FechaEmision,
HH.FechaCancela,HH.TipoCambio,		
DebeS=
	case when YY.diferencia<0 then round(abs(YY.Diferencia),2) else 0 end,
HaberS=
	case when YY.diferencia>0 then Round(YY.Diferencia,2) else 0 end,
DebeD=
	case when YY.diferencia<0 
		then case when YY.diferencia<0 then round(abs(YY.Diferencia)/HH.TipoCambio,2) else 0 end
		else case when YY.diferencia<0 then round(abs(YY.Diferencia),2) else 0 end
	end,
HaberD=
	case when HH.TipoCambio>0 
		then	case when YY.diferencia>0 then round(YY.Diferencia/HH.TipoCambio,2) else 0 end
		else  case when YY.diferencia>0 then round(YY.Diferencia,2) else 0 end
	end,
HH.ImporteSol,
HH.ImporteDolar,
HH.ImporteSoles                                         

from                                                 
	(select cabrec_numrecibo,Diferencia=Round(sum(DebeS),2)-Round(Sum(HaberS),2) from 
		(select * from ##tmpCuentaEgreso' +@NombrePC+ ' union all
 			select * from ##tmpCuentaCaja' +@NombrePC+ ')as ZZ
			group by cabrec_numrecibo)as YY,
	##tmpCuentaCaja' +@NombrePC+ ' HH
where YY.cabrec_numrecibo=HH.cabrec_numrecibo and YY.Diferencia<>0 ) as WW
order by 1,2'

exec(@CadDifCambio) 
set nocount off

--exec te_GeneraTempContable 'Ventas_Prueba','Contaprueba','776101','976101','Desarrollo3','12','2002','E'


--select * from ##tmpAsientosContaDesarrollo3 order by 1
--select * from ##tmpAsientosContaDesarrollo3 where cuenta not in (select cuenta from contaprueba..ct_cuenta)

--select * from ct_tipocambio order by 1 desc
--select * from ventas_prueba.dbo.cp_cargo where cargonumdoc like '%12357%'