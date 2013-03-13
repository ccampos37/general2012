select * from te_detallerecibos where cabrec_numrecibo='101062' 


select * from vt_abono where documentoabono+abononumdoc in

delete from vt_abono where documentoabono+abononumdoc in
select * from vt_abono where documentoabono+abononumdoc in
(select detrec_tipodoc_concepto+detrec_numdocumento from te_detallerecibos where cabrec_numrecibo='101005')

select * from vt_cargo 

update vt_cargo set cargoapeimppag=0,cargoapeflgcan=0
select * from vt_cargo
where documentocargo+cargonumdoc in
	(select detrec_tipodoc_concepto+detrec_numdocumento from te_detallerecibos where cabrec_numrecibo='101005')

--documentoabono+abononumdoc+abonocancli

select * from te_cabecerarecibos where cabrec_numrecibo='101006'
select * from te_detallerecibos where cabrec_numrecibo='101006'
update te_cabecerarecibos set cabrec_estadoreg=0 where cabrec_numrecibo='101006'
update te_detallerecibos set detrec_estadoreg=0 where cabrec_numrecibo='101006'


