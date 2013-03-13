/*
select * from contaprueba.dbo.ct_detcomprob2002
cabcomprobmes cabcomprobnumero subasientocodigo analiticocodigo asientocodigo detcomprobitem 
monedacodigo centrocostocodigo documentocodigo operacioncodigo cuentacodigo         
detcomprobnumdocumento detcomprobfechaemision  detcomprobfechavencimiento detcomprobglosa 
detcomprobdebe detcomprobhaber  detcomprobusshaber  detcomprobussdebe  detcomprobtipocambio
detcomprobruc detcomprobauto detcomprobformacambio detcomprobajusteuser plantillaasientoinafecto 
tipdocref detcomprobnumref  detcomprobconci detcomprobnlibro detcomprobfecharef 
*/


Declare @CadCtasCaja varchar(5000)
Declare @Base varchar(50)
Declare @BaseConta varchar(50)
Declare @NombrePC varchar(50)


set @Base='Ventas_Prueba'
set @BaseConta='Contaprueba'
set @NombrePC='Desarrollo3'



if exists (Select name from tempdb.dbo.sysobjects where name='##tmpCuentaCaja' +@NombrePC)
  exec('drop table ##tmpCuentaCaja' +@NombrePC)

set @CadCtasCaja='
select
	 	cabrec_numrecibo,
		detrec_item=
		isnull((select max(detrec_item) from ##tmpEgreso' +@NombrePC+ ' A where cabrec_numrecibo=TT.cabrec_numrecibo),0)+1,
		detrec_monedacancela as MonedaCodigo,
		detrec_tipodoc_concepto,
		detrec_numdocumento,
		clientecodigo,
		cuenta,detrec_observacion,
		detrec_fechacancela as FechaCancela,tccancela as TipoCambio,
		DebeS=0.000,
		HaberS=
			case when detrec_monedacancela=''02'' 
				then	Round(ImporteDolar*tccancela,2)
				else	Round(ImporteSoles,2)
			end,
		DebeD=0.00,
		HaberD=Round(ImporteDolar,2),
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
		group by bb.cabrec_numrecibo) as ZZ,
	[' +@Base+ '].dbo.te_detallerecibos YY	
where 
	ZZ.cabrec_numrecibo=YY.cabrec_numrecibo) as XX ) as TT
order by 1'

exec(@CadCtasCaja)

/*
if exists(select name from tempdb.dbo.sysobjects where name='##tmpEgresoDesarrollo3')
	exec ('drop table ##tmpEgresoDesarrollo3')
*/


/*
select * from dbo.te_cuentabancos
--select * from dbo.te_detallerecibos

update te_detallerecibos set detrec_numctacte='011-350-0100008495-62'
--select * from te_detallerecibos
where detrec_tipocajabanco='B' and 
		detrec_cajabanco1='02' and
		detrec_monedacancela='01' 
011-350-0100008495-62
*/
