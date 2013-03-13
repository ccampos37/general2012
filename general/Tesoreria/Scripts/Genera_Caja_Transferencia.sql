--select * into XXXX from 

'select * into ##tmpAsientosContaTransf' +@NombrePC+ ' from 

(select * from ##tmpCuentaCajaTransfDesarrollo3 union all
		select * from 
(select distinct
	HH.cabrec_numreciboegreso,
	HH.cabrec_numrecibo,
   HH.cabrec_ingsal,
detrec_item=(select max(detrec_item) from ##tmpCuentaCajaTransfDesarrollo3 A where cabrec_numreciboegreso=YY.cabrec_numreciboegreso)+1  ,
HH.MonedaCodigo,HH.detrec_tipodoc_concepto, 
HH.detrec_numdocumento,HH.clientecodigo,   
cuenta= 
  case when YY.diferencia>0 then '776101' else '976101' end,
HH.detrec_observacion,
HH.FechaEmision,
HH.FechaCancela,HH.TipoCambio,		
DebeS=
	case when YY.diferencia<0 then round(abs(YY.Diferencia),2) else 0 end,
HaberS=
	case when YY.diferencia>0 then Round(abs(YY.Diferencia),2) else 0 end,
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
		(select * from ##tmpCuentaCajaTransfDesarrollo3)as ZZ
			group by cabrec_numreciboegreso
         having Round(sum(DebeS),2)-Round(Sum(HaberS),2)<>0 )as YY,
	##tmpCuentaCajaTransfDesarrollo3 HH
where YY.Diferencia<>0 and  YY.cabrec_numreciboegreso=HH.cabrec_numreciboegreso and HH.cabrec_ingsal='E') as WW ) as ZZ
order by 1

--select * from 	##tmpCuentaCajaTransfDesarrollo3 HH