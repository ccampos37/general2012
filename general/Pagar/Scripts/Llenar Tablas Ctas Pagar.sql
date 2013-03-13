insert cp_tipodocumento 	SELECT * from cc_tipodocumento
insert cp_tipoplanilla 		select * from cc_tipoplanilla
insert cp_conceptos 		select * from cc_conceptos

select * into cp_proveedor 				from 	vt_cliente
select * into cp_proveedordireccion 	from 	vt_clientedireccion
select * into cp_cargo 					from 	vt_cargo
select * into cp_abono 					from 	vt_abono
select * into Tempo_proveedordireccion 	from 	Tempo_clientedireccion
select * into cp_oficina  				from 	vt_vendedor

select * into cp_negocio 				from 	vt_negocio

select * into cp_parametros				from	gr_empresa