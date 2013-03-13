select stalma,stcodigo,stskdis,x=right(rtrim(id_empresa),2) ,producto_id ,stock,id_empresa
from planta10.dbo.productoempresa  
full join stkart
on right(rtrim(id_empresa),2)=stalma and producto_id=stcodigo
inner join tabalm on right(rtrim(id_empresa),2)=taalma
where isnull(stskdis,0)<>isnull(stock,0) 



select * from stkart where stskdis<>0 and stalma='01'
select * from planta10.dbo.productoempresa where id_empresa='001' and isnull(stock,0) <>0

/*
select * from planta10.dbo.productoempresa where stock> 0 and right(rtrim(id_empresa),2)='01' order by 2
select * from planta_casma.dbo.stkart where stalma='01' order by 2

select * from stkart where stcodigo='10586' 
select * from planta10.dbo.productoempresa where producto_id='10586'
select a.*,b.* from movalmcab a inner join movalmdet b on caalma+catd+canumdoc=dealma+detd+denumdoc where decodigo='10586'

select * into xx_movalmdet from movalmdet

update movalmdet
set deprecio=yy.precio
from movalmdet a,
( select dealma,detd,denumdoc,deitem,decodigo,deprecio,precio
from 
( select b.* from movalmcab a inner join movalmdet b on caalma+catd+canumdoc=dealma+detd+denumdoc where catipmov='I'
) as z
inner join
( select a.id_empresa,a.producto_id,precio=preciounitariocompra from planta10.dbo.kardex a
   inner join 
      (select id_empresa,producto_id,item= max(itemkardex) from planta10.dbo.kardex 
             where cantidadingreso > 0
             group by id_empresa,producto_id ) as z
   on a.id_empresa=z.id_empresa and a.producto_id=z.producto_id and itemkardex=item
) as zz
on z.dealma=right(rtrim(id_empresa),2) and z.decodigo=zz.producto_id
) as yy
where a.dealma=yy.dealma and a.detd=yy.detd and a.denumdoc=yy.denumdoc and a.deitem=yy.deitem
and a.detd='NI' and a.deprecio=0





select * from planta10.dbo.kardex
inner join planta10.dbo.productoempresa on stalma=right(rtrim(id_empresa),2) and stcodigo=producto_id 
inner join tabalm on stalma=taalma
*/ 