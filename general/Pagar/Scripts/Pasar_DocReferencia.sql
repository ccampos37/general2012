select * from cp_cargo


replicate('0',8-len(cabprovinumdoc))

select left(cabprovinumdoc,3) +  substring(cabprovinumdoc,5,len(cabprovinumdoc)-4)
from co_cabprovi2003 --where cabprovinumdoc

select 
	cabprovinumdoc, 
	left(cabprovinumdoc,patindex('%-%',cabprovinumdoc)-1),
	right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc)),
	left(cabprovinumdoc,patindex('%-%',cabprovinumdoc)-1)+replicate('0',8-len(right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))))+right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))
from co_cabprovi2002 --where cabprovinumdoc


--select * from co_cabprovi2003  where cabprovifchdoc<>cabprovifchven 
-- where replicate('0',6-len(cabprovinumero))+ cast(cabprovinumero as varchar(6))
-- 		in (select abononumplanilla from cp_cargo)
cabprovitipdocref cabprovinref


--select * into bk_cp_cargo from cp_cargo a,
update cp_cargo set
	cargoapetiporefe=zz.cabprovitipdocref,
	cargoapenrorefe=rtrim(ltrim(left(zz.cabprovinref,11)))
from 
	cp_cargo a,
(select a.cabprovitipdocref,a.cabprovinref,
	b.cargonumdoc,
	b.documentocargo,
	b.clientecodigo,
	b.abononumplanilla,
	numdoc=left(cabprovinumdoc,patindex('%-%',cabprovinumdoc)-1)+replicate('0',8-len(right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))))+right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))
	from co_cabprovi2003 a,
			cp_cargo b
where 
	left(cabprovinumdoc,patindex('%-%',cabprovinumdoc)-1)+replicate('0',8-len(right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))))+right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))=b.cargonumdoc and
	b.clientecodigo=a.proveedorcodigo and
	b.documentocargo=a.documetocodigo and 
   (a.cabprovitipdocref<>'' and cabprovinref<>'')) as ZZ
where 
 	a.clientecodigo=zz.clientecodigo and a.cargonumdoc=zz.cargonumdoc and a.documentocargo=zz.documentocargo


/*Actualizar Fecha Vencimiento*/
--update cp_cargo set
--	cargoapefecvct=zz.cabprovifchven
select a.* 
from 
	cp_cargo a,
(select a.cabprovifchdoc,
	a.cabprovifchven,
	b.cargonumdoc,
	b.documentocargo,
	b.clientecodigo,
	b.abononumplanilla,
	numdoc=left(cabprovinumdoc,patindex('%-%',cabprovinumdoc)-1)+replicate('0',8-len(right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))))+right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))
	from co_cabprovi2003 a,
			cp_cargo b
where 
	left(cabprovinumdoc,patindex('%-%',cabprovinumdoc)-1)+replicate('0',8-len(right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))))+right(cabprovinumdoc,len(cabprovinumdoc)-patindex('%-%',cabprovinumdoc))=b.cargonumdoc and
	b.clientecodigo=a.proveedorcodigo and
	b.documentocargo=a.documetocodigo) as ZZ
where 
 	a.clientecodigo=zz.clientecodigo and a.cargonumdoc=zz.cargonumdoc and a.documentocargo=zz.documentocargo