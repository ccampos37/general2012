SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


alter   Procedure te_GeneraTempContableTransf(

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


if exists (Select name from tempdb.dbo.sysobjects where name='##tmpCuentaCajaTransf' +@NombrePC)
  exec('drop table ##tmpCuentaCajaTransf' +@NombrePC)

set @CadCtasCaja='
select
		cabrec_numreciboegreso,
	 	cabrec_numrecibo,
      cabrec_ingsal,
		detrec_item=
		isnull((select max(detrec_item) from ##tmpEgresoDesarrollo3 A where cabrec_numrecibo=TT.cabrec_numrecibo),0)+1,
		detrec_monedacancela as MonedaCodigo,
		detrec_tipodoc_concepto,
		detrec_numdocumento,
		clientecodigo,
		cuenta,detrec_observacion,
		detrec_fechacancela as FechaEmision,
		detrec_fechacancela as FechaCancela,tccancela as TipoCambio,
     	DebeS=
			case when cabrec_ingsal=''I''
				then 0.00
				else	
					case when detrec_monedacancela=''02'' 
						then	Cast(Round(ImporteDolar*tccancela,2) as Numeric(15,2))
						else	Cast(Round(ImporteSoles,2) as Numeric(15,2))
					end
			end,
		HaberS=
			case when cabrec_ingsal=''I''
				then 
					case when detrec_monedacancela=''02'' 
						then	Cast(Round(ImporteDolar*tccancela,2) as Numeric(15,2))
						else	Cast(Round(ImporteSoles,2) as Numeric(15,2))
					end
				else
					0.00
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

into ##tmpCuentaCajaTransfDesarrollo3

from
(Select XX.*,  
		tccancela=
			case when XX.detrec_monedacancela=''02'' 
					then	isnull((select tipocambioventa from [Contaprueba].dbo.ct_tipocambio as M where M.tipocambiofecha=XX.detrec_fechacancela),0)
					else	1
			end
from
(select 	distinct ZZ.cabrec_numreciboegreso,ZZ.cabrec_numrecibo,
			ZZ.cabrec_ingsal, 
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
									(select C.cajacuentasoles from [Ventas_Prueba].dbo.te_codigocaja C where C.cajacodigo=YY.detrec_cajabanco1)
								else
									(select C.cajacuentadolares from [Ventas_Prueba].dbo.te_codigocaja C where C.cajacodigo=YY.detrec_cajabanco1)
							end
					else
						(select cbanco_cuenta from [Ventas_Prueba].dbo.te_cuentabancos
							where cbanco_codigo=YY.detrec_cajabanco1 and monedacodigo=YY.detrec_monedacancela and cbanco_numero=YY.detrec_numctacte)
				end,
			detrec_tdqc=(select top 1 A.detrec_tdqc from [Ventas_Prueba].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_ndqc=(select top 1 A.detrec_ndqc from [Ventas_Prueba].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_observacion=(select top 1 A.detrec_observacion from [Ventas_Prueba].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_fechacancela=(select top 1 A.detrec_fechacancela from [Ventas_Prueba].dbo.te_detallerecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo),
			detrec_tipodoc_concepto=''00'',
			detrec_numdocumento=cast('''' as varchar(50)),
			clientecodigo=(select top 1 clientecodigo from [Ventas_Prueba].dbo.te_cabecerarecibos A where A.cabrec_numrecibo=YY.cabrec_numrecibo)
from
	(select 	CC.cabrec_numreciboegreso,
				bb.cabrec_numrecibo,
				cc.cabrec_ingsal,
				ImporteSoles=sum(bb.detrec_importesoles),
				ImporteDolar=sum(bb.detrec_importedolares) 
		from [Ventas_Prueba].dbo.te_detallerecibos bb,
			  [Ventas_Prueba].dbo.te_cabecerarecibos cc	
			
		where  bb.cabrec_numrecibo=cc.cabrec_numrecibo and 
				 isnull(cc.cabrec_transferenciaautomatico,0)=''1''	and 
			month(bb.detrec_fechacancela)=' +@MesProceso+' and year(bb.detrec_fechacancela)=' +@AnnoProceso+ ' 
		group by CC.cabrec_numreciboegreso,bb.cabrec_numrecibo,cc.cabrec_ingsal ) as ZZ,
	[Ventas_Prueba].dbo.te_detallerecibos YY	
where 
	ZZ.cabrec_numrecibo=YY.cabrec_numrecibo) as XX ) as TT
order by 1'

exec(@CadCtasCaja)


/*

if exists (Select name from tempdb.dbo.sysobjects where name='##tmpAsientosConta' +@NombrePC)
  exec('drop table ##tmpAsientosConta' +@NombrePC)

set @CadDifCambio='


select * into ##tmpAsientosConta' +@NombrePC+ ' 
	select * from ##tmpCuentaCajaTransf' +@NombrePC+ ' union all
		select * from 
(select YY.cabrec_numreciboegreso,
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
	(select cabrec_numreciboegreso,Diferencia=Round(sum(DebeS),2)-Round(Sum(HaberS),2) from 
		(select * from ##tmpCuentaCajaTransf' +@NombrePC+ ')as ZZ
			group by cabrec_numreciboegreso)as YY,
	##tmpCuentaCajaTransf' +@NombrePC+ ' HH
where YY.Diferencia<>0 ) as WW
order by 1'

print(@CadDifCambio) 
*/


set nocount off

--exec te_GeneraTempContableTransf 'Ventas_Prueba','Contaprueba','776101','976101','Desarrollo3','12','2002','%'


--select * from ##tmpAsientosContaDesarrollo3 order by 1
--select * from ##tmpAsientosContaDesarrollo3 where cuenta not in (select cuenta from contaprueba..ct_cuenta)

--select * from ##tmpCuentaCajaTransfDesarrollo3

--select * from ct_tipocambio order by 1 desc
--select * from ventas_prueba.dbo.cp_cargo where cargonumdoc like '%12357%'


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

