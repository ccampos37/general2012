select B.cuentacodigo,
  EneD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=1),
  EneH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=1),	

  FebD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=2),
  FebH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=2),	

  MarD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=3),
  MarH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=3),	

  AbrD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=4),
  AbrH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=4),	

  MayD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=5),
  MayH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=5),	

  JunD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=6),
  JunH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=6),	

  JulD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=7),
  JulH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=7),	

  AgoD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=8),
  AgoH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=8),	

  SetD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=9),
  SetH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=9),

  OctD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=10),
  OctH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=10),	

  NovD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=11),
  NovH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=11),	

  DicD=(select sum(A.detcomprobdebe) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=12),
  DicH=(select sum(A.detcomprobhaber) from ct_detcomprob2002 A 
			where A.cabcomprobmes=B.cabcomprobmes and A.cuentacodigo=B.cuentacodigo and A.cabcomprobmes=12)	

from ct_detcomprob2002 B
where left(cuentacodigo,2) in ('92','94','95','97')
group by B.cabcomprobmes,B.cuentacodigo
order by B.cabcomprobmes,B.cuentacodigo