/*
select cargonumdoc,left(cargonumdoc,3),
 ltrim(rtrim(replicate('0',8-len(substring(cargonumdoc,4,len(cargonumdoc)-3))))) as Ceros,
 substring(cargonumdoc,4,len(cargonumdoc)-3)
   from cp_cargo where len(cargonumdoc)<>11 and documentocargo+cargonumdoc+clientecodigo not in (select documentoabono+abononumdoc+abonocancli from cp_abono)
*/

select clientecodigo,documentocargo,cargonumdoc,left(cargonumdoc,3)+
 ltrim(rtrim(replicate('0',8-len(substring(cargonumdoc,4,len(cargonumdoc)-3)))))+substring(cargonumdoc,4,len(cargonumdoc)-3) as NuevoNum
--  into ListaDocNuevoNumero
   from cp_cargo where len(cargonumdoc)<>11 and documentocargo+cargonumdoc+clientecodigo not in (select documentoabono+abononumdoc+abonocancli from cp_abono)


--select count(*),documentocargo,cargonumdoc,clientecodigo
--from
(select a.*--,b.*
from cp_cargo a, ListaDocNuevoNumero b
where a.clientecodigo=b.clientecodigo and 
		a.documentocargo=b.documentocargo and
		a.cargonumdoc=b.cargonumdoc) as ZZ
group by documentocargo,cargonumdoc,clientecodigo
having count(*)>1

--            documentocargo cargonumdoc clientecodigo 
------------- -------------- ----------- ------------- 
--4           03             H00388168   20131380951
--4           10             017061323   10073161121

update cp_cargo set cargonumdoc=b.nuevonum 
--select a.*,b.*
from cp_cargo a, ListaDocNuevoNumero b
where a.clientecodigo=b.clientecodigo and 
		a.documentocargo=b.documentocargo and
		a.cargonumdoc=b.cargonumdoc


select * from cp_cargo where documentocargo='10' and cargonumdoc like '017%061323'
select * from cp_cargo where documentocargo='01' and cargonumdoc like '001%94609'
--select * from ListaDocNuevoNumero where documentocargo='01' and cargonumdoc like '001%94609'

select * into bk_cp_cargo from cp_cargo 