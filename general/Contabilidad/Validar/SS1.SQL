select cabcomprobmes,cabcomprobnumero,debe=sum(case when cuentacodigo like '2%' then detcomprobdebe else detcomprobdebe*-1 end),
	haber=sum(case when cuentacodigo like '2%' then detcomprobhaber else detcomprobhaber*-1 end)
    from ct_detcomprob2002 
   where cuentacodigo like '20%' or cuentacodigo like '24%' or cuentacodigo like '26%' or cuentacodigo like '60%'
 group by cabcomprobmes,cabcomprobnumero
 having sum(case when cuentacodigo like '2%' then detcomprobdebe else detcomprobdebe*-1 end)-sum(case when cuentacodigo like '2%' then detcomprobhaber else detcomprobhaber*-1 end)<>0
 order by 1