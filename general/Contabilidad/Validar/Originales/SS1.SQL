select cabcomprobmes,cabcomprobnumero,debe=sum(case when cuentacodigo like '7%' then detcomprobdebe else detcomprobdebe*-1 end),
	haber=sum(case when cuentacodigo like '7%' then detcomprobhaber else detcomprobhaber*-1 end)
    from ct_detcomprob2002 
   where cuentacodigo like '62%' or cuentacodigo like '63%' or cuentacodigo like '64%' or cuentacodigo like '65%' or cuentacodigo like '67%' or cuentacodigo like '68%' 
       or cuentacodigo like '79%'
 group by cabcomprobmes,cabcomprobnumero
 having sum(case when cuentacodigo like '7%' then detcomprobdebe else detcomprobdebe*-1 end)+sum(case when cuentacodigo like '7%' then detcomprobhaber else detcomprobhaber*-1 end)<>0
 order by 1