select a.*,b.* /*,c.bancodescripcion*/ from 
	[ventas_prueba].dbo.cp_cargo a, 
	[ventas_prueba].dbo.cp_tipoplanilla b,
	gr_banco c
where 
	a.abonotipoplanilla=b.tplanillacodigo and
	(isnull(b.tplanillacanjes,0)=1 or isnull(b.tplanillarenovar,0)=1) and
	isnull(a.cargoapeflgcan,0)<>1 and
	a.bancocodigo=c.bancocodigo and 
	isnull(a.bancocodigo,0) like '%' 

--select * from cp_cargo
--select * from cp_tipoplanilla
--select * from gr_banco

/*
select * from vt_cargo where clientecodigo='65' and documentocargo='62' and cargonumdoc='00000099999'
select * from vt_abono
*/