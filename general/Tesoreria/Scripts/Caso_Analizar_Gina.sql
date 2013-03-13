select * from cp_cargo where cargonumdoc like '001%53835' and documentocargo='02'
update cp_cargo set cargoapeflgreg='1' where cargonumdoc like '001%53835' and documentocargo='02'



select * from cp_cargo 
	where cargonumdoc like '001%326' and clientecodigo='20300030981'
select * from cp_abono
	where abononumdoc like '001%326' and abonocancli='20300030981'

select * from co_cabprovi2003 where cabprovinumero='228'

select * from cp_abono where abononumplanilla='202059'

update cp_abono set abonocanmoneda='01' where abononumplanilla='202059' and abonocanimpcan<>abonocanimpsol 