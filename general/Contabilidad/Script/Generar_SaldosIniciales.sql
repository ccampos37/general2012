select * from ct_saldos2002 where saldodebe00<>0 or saldohaber00<>0
select * from ct_saldos2003 where saldodebe00<>0 or saldohaber00<>0


select cuentacodigo,saldoacumussdebe12,saldoacumusshaber12 from ct_saldos2002
	where (saldoacumussdebe12<>0 or saldoacumusshaber12<>0)
			and cuentacodigo not in (select cuentacodigo from ct_saldos2003)

update ct_saldos2003
	set saldodebe00=0,
		 saldohaber00=0


update ct_saldos2003
	set saldodebe00=case when b.saldoacumdebe12-b.saldoacumhaber12>0 then abs(b.saldoacumdebe12-b.saldoacumhaber12) else 0 end,
		 saldohaber00=case when b.saldoacumdebe12-b.saldoacumhaber12<0 then abs(b.saldoacumdebe12-b.saldoacumhaber12) else 0 end
--select b.* 
from ct_saldos2003 a, ct_saldos2002 b
where a.cuentacodigo=b.cuentacodigo



insert contaprueba.dbo.ct_saldos2003
select * from contaprueba.dbo.ct_cuenta 
	where cuentacodigo not in (select cuentacodigo from contaprueba.dbo.ct_saldos2003)
			and len(cuentacodigo)=6

update ct_saldos2003
	set saldodebe00=case when b.saldoacumdebe12-b.saldoacumhaber12>0 then abs(b.saldoacumdebe12-b.saldoacumhaber12) else 0 end,
		 saldohaber00=case when b.saldoacumdebe12-b.saldoacumhaber12<0 then abs(b.saldoacumdebe12-b.saldoacumhaber12) else 0 end
--select b.* 
from contaprueba.dbo.ct_saldos2003 a, contaprueba.dbo.ct_saldos2002 b
where 
	a.cuentacodigo=b.cuentacodigo and left(a.cuentacodigo,2)<='59'