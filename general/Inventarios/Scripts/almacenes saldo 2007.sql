select * from
(
select dealma,denumdoc,decodigo,decantid from planta_casma.dbo.movalmdet aa
inner join planta_casma.dbo.movalmcab b 
on dealma=caalma and detd=catd and denumdoc=canumdoc
where cafecdoc>='01/01/2008' and detd='NI'
) as a
right join 
(
select empresa=right(rtrim(b.id_empresa),2),doc=right('00000000000'+rtrim(b.registroingreso_id),11),
codigo=producto_id,cantidadingresada from planta10.dbo.detregistroingreso a 
inner join planta10.dbo.registroingreso b on a.registroingreso_id=b.registroingreso_id
where b.id_empresa in  ('001','002') and b.fecharecepcion > ='01/01/2008'  
) as b
on a.dealma=b.empresa and a.denumdoc=b.doc and a.decodigo=b.codigo
 
/*
drop table yy_saldo2007 
select z.dealma,z.decodigo,saldo=z.stock+z.ingresos+z.salidas
into yy_saldo2007
from
(

select dealma,decodigo,stock,
ingresos= sum(case when b.catipmov='I' then decantid*-1 else 0 end),
salidas=sum(case when b.catipmov='I' then 0 else decantid end)
 from planta_casma.dbo.movalmdet a
inner join planta_casma.dbo.movalmcab b on dealma=caalma and detd=catd and denumdoc=canumdoc
inner  join planta10.dbo.productoempresa c on '0'+dealma=id_empresa and decodigo=producto_id  
where id_empresa in ('001','002')  and cafecdoc>='01/01/2008'
group by dealma,decodigo,stock
) as z 
where z.stock+z.ingresos+z.salidas>0
*/

select z.dealma,z.decodigo,d.adescri,saldo=z.saldo+cantidad,actual=stock from
(
select a.dealma,a.decodigo,saldo,
cantidad=sum(case when c.catipmov='I' then b.decantid else b.decantid*-1 end)
from planta_casma.dbo.yy_saldo2007 a
inner join planta_casma.dbo.movalmdet b on a.dealma=b.dealma and a.decodigo=b.decodigo
inner join planta_casma.dbo.movalmcab c on b.dealma=c.caalma and b.detd=c.catd and b.denumdoc=c.canumdoc
where c.cafecdoc>='01/01/2008'
group by a.dealma,a.decodigo,saldo
) as z
inner join planta_casma.dbo.maeart d on z.decodigo=d.acodigo 
inner  join planta10.dbo.productoempresa c on '0'+z.dealma=id_empresa and z.decodigo=producto_id  
where saldo+cantidad<>stock

select dealma,decodigo,adescri,saldo from yy_saldo2007 z
inner join planta_casma.dbo.maeart d on z.decodigo=d.acodigo order by 1,2
