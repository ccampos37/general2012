select sum(dif)
from
(select 
   	cabcomprobnumero,
    	Dif=sum(detcomprobdebe)-sum(detcomprobhaber) 
  	from dbo.ct_detcomprob2002
	where cabcomprobmes=12
  group by cabcomprobnumero
  having sum(detcomprobdebe)-sum(detcomprobhaber)<>0 ) as ZZ
order by 1 asc