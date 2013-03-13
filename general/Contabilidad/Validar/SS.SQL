/*Clase 9 con la Clase 6*/
select cabcomprobmes,cabcomprobnumero,debe=sum(case when cuentacodigo like '9%' then detcomprobdebe else detcomprobdebe*-1 end),
	haber=sum(case when cuentacodigo like '9%' then detcomprobhaber else detcomprobhaber*-1 end)
    from ct_detcomprob2003 
   where cuentacodigo like '62%' or cuentacodigo like '63%' or cuentacodigo like '64%' or cuentacodigo like '65%' or cuentacodigo like '67%' or cuentacodigo like '68%' 
       or cuentacodigo like '92%' or cuentacodigo like '94%' or cuentacodigo like '95%' or cuentacodigo like '97%'
 group by cabcomprobmes,cabcomprobnumero
 having sum(case when cuentacodigo like '9%' then detcomprobdebe else detcomprobdebe*-1 end)-sum(case when cuentacodigo like '9%' then detcomprobhaber else detcomprobhaber*-1 end)<>0
 order by 1