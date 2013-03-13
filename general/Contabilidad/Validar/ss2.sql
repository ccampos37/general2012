select cabcomprobmes,cabcomprobnumero,debe=sum(case when cuentacodigo like '60%' then detcomprobdebe else detcomprobdebe*-1 end),
	haber=sum(case when cuentacodigo like '60%' then detcomprobhaber else detcomprobhaber*-1 end)
    from ct_detcomprob2002 
   where cuentacodigo like '60%' or cuentacodigo like '61%' 
 group by cabcomprobmes,cabcomprobnumero
 having sum(case when cuentacodigo like '60%' then detcomprobdebe else detcomprobdebe*-1 end)+sum(case when cuentacodigo like '60%' then detcomprobhaber else detcomprobhaber*-1 end)<>0
 order by 1