117--Filas
'776101' --Ganancia
'976101' --Perdida
use contaprueba

 Select
 cabcomprobmes,
 cabcomprobnumero,
 subasientocodigo,
 analiticocodigo='00',
 asientocodigo,
 detcomprobitem='00099',
 monedacodigo='01',
 centrocostocodigo='00',
 documentocodigo,
 operacioncodigo='00',
 cuentacodigo=B.cuenta, 
 detcomprobnumdocumento,
 detcomprobfechaemision,
 detcomprobfechavencimiento,
 detcomprobglosa,
 detcomprobdebe=case when B.dif > 0 then 0 else abs(B.dif) end ,
 detcomprobhaber=case when B.dif > 0 then abs(B.dif) else 0 end, 
 detcomprobusshaber=0, 
 detcomprobussdebe=0,
 detcomprobtipocambio=0,
 detcomprobruc,
 detcomprobauto, 
 detcomprobformacambio,
 detcomprobajusteuser,
 plantillaasientoinafecto=0, 
 tipdocref,
 detcomprobnumref,
 detcomprobconci,
 detcomprobnlibro,
 detcomprobfecharef=getdate()
 from  ct_detcomprob2003 A, 
 (select 
    cabcomprobnumero,
    cuenta=case when (sum(detcomprobdebe)-sum(detcomprobhaber)) > 0 then '776101' else '976101' end, 
	Debe=sum(detcomprobdebe),Haber=sum(detcomprobhaber),
    Dif=sum(detcomprobdebe)-sum(detcomprobhaber) 
  from dbo.ct_detcomprob2003
  where cabcomprobmes=1 and monedacodigo='02'
  group by cabcomprobnumero
  having sum(detcomprobdebe)-sum(detcomprobhaber)<>0 ) as B
 Where 
 A.cabcomprobnumero=B.cabcomprobnumero and 
 A.cuentacodigo like '70%' and cabcomprobmes=1
 
        
 


