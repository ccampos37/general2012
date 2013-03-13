Declare @CadDifCambio varchar(4000)
Declare @CtaGanCambio varchar(20)
Declare @CtaPerCambio varchar(20)
Declare @NombrePC varchar(50)


set @CtaGanCambio='776101'
set @CtaPerCambio='976101'
set @NombrePC='Desarrollo3'


set @CadDifCambio='
select * from ##tmpCuentaEgreso' +@NombrePC+ ' union all
	select * from ##tmpCuentaCaja' +@NombrePC+ ' union all
		select * from 
(select YY.cabrec_numrecibo,
detrec_item=(select max(detrec_item) from ##tmpCuentaCaja' +@NombrePC+ ' A where cabrec_numrecibo=YY.cabrec_numrecibo)+1  ,
HH.MonedaCodigo,HH.detrec_tipodoc_concepto, 
HH.detrec_numdocumento,HH.clientecodigo,   
cuenta= 
  case when YY.diferencia>0 then ''' +@CtaGanCambio+  ''' else ''' +@CtaPerCambio+ ''' end,
HH.detrec_observacion,
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