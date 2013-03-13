
select sum(di) from (

--update ct_detcomprob2003 set detcomprobhaber=Z.nigv
select ct_detcomprob2003.cabcomprobnumero,ct_detcomprob2003.cuentacodigo,Z.nigv 
from ct_detcomprob2003,
(select detcomprobnumdocumento , cabcomprobnumero,di =sum(round(monto*.18,2))-sum(igv),nigv=sum(round(monto*.18,2)),ig=sum(igv),Mo= sum(monto)
from (
select detcomprobnumdocumento , cabcomprobnumero,monedacodigo,
monto = round(case when cuentacodigo='701101' then detcomprobhaber else 0.00 end,2) ,
igv   = round(case when cuentacodigo='401110' then detcomprobhaber else 0.00 end,2)
from ct_detcomprob2003
where cabcomprobmes=1 and cabcomprobnumero like '0107%' and asientocodigo like '072%' and  documentocodigo ='03' and
      detcomprobhaber > 0 and left(detcomprobnumdocumento,3)  in('008')) as X
group by detcomprobnumdocumento,cabcomprobnumero
having sum(round(monto*.18,2))-sum(igv)< 0.00 ) as z 
where ct_detcomprob2003.cabcomprobnumero=Z.cabcomprobnumero and ct_detcomprob2003.cuentacodigo='401110'


select * from descuadreigv



update ct_detcomprob2003 set detcomprobdebe=Z.nmonto

--select ct_detcomprob2003.cabcomprobnumero,ct_detcomprob2003.cuentacodigo,Z.nmonto,z.ne,detcomprobdebe
from ct_detcomprob2003,
(select cabcomprobmes , cabcomprobnumero,di =sum(monto)-sum(neto),nmonto=sum(monto), ne=sum(neto)
from (
select cabcomprobmes ,cabcomprobnumero,cuentacodigo,monedacodigo,
monto = round(case when cuentacodigo='701101' or cuentacodigo='401110' then detcomprobhaber else 0.00 end,2) ,
neto  = round(case when cuentacodigo like '12%' then detcomprobdebe else 0.00 end,2)
from ct_detcomprob2003
where cabcomprobmes=1 and cabcomprobnumero like '0107%' and asientocodigo like '072%' and  documentocodigo ='03' 
      and left(detcomprobnumdocumento,3)  in('003')) as X
group by cabcomprobmes ,cabcomprobnumero 
having sum(neto)-sum(monto)> 0 ) as z 
where ct_detcomprob2003.cabcomprobnumero=Z.cabcomprobnumero and ct_detcomprob2003.cuentacodigo like '121%'


