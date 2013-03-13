/*
201942 con fecha 19/03/2003
*/

select * from te_cabecerarecibos where cabrec_numrecibo='201942'
select * from te_detallerecibos where cabrec_numrecibo='201942'

update te_detallerecibos set detrec_fechacancela='19/03/2003'
where cabrec_numrecibo='201942'

update te_cabecerarecibos set cabrec_fechadocumento='19/03/2003'
where cabrec_numrecibo='201942'

select * from te_cabecerarecibos where cabrec_numrecibo='200572'

select * from cp_abono where abononumdoc in (select detrec_numdocumento from te_detallerecibos where cabrec_numrecibo='201258')     

update cp_abono set abonocanfecpla='19/03/2003',abonocanfecpro='19/03/2003',abonocanfecan='19/03/2003'
--select * from cp_abono
where cp_abono.abononumplanilla='201942'
	(select detrec_numdocumento from te_detallerecibos where cabrec_numrecibo='204076')     

/*
update cp_abono set abonocanfecpla='28/02/2003',abonocanfecpro='28/02/2003',abonocanfecan='28/02/2003'
--select * from cp_abono
where cp_abono.abononumdoc in 
	(select detrec_numdocumento from te_detallerecibos where cabrec_numrecibo='202346')     
*/

update cp_abono set abonocanfecpla='16/04/2003',abonocanfecpro='16/04/2003',abonocanfecan='16/04/2003'
--select * from cp_abono
where abononumplanilla='203168' 


update cp_abono set abonocanfecpla='02/11/2002',abonocanfecpro='02/11/2002',abonocanfecan='02/11/2002'
where abononumplanilla='000132' and abononumdoc='00100000005'