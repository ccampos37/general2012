use camtex_tj


select * from vt_cargo
select * from vt_abono

select * from cp_cargo

select * from cp_abono

select *  from te_cabecerarecibos order by cabrec_numreciboegreso desc
select *  from te_detallerecibos order by cabrec_numrecibo desc



select cabrec_numrecibo from dbo.te_cabecerarecibos_bk --where cabrec_numrecibo='202290'
--select * from dbo.te_detallerecibos_bk --where cabrec_numrecibo='202290'
  where isnull(cabrec_numreciboegreso,0)='' and isnull(cabrec_estadoreg,0)<>1 
   

update dbo.te_cabecerarecibos set  cabrec_numrecibo




select cabrec_numrecibo,count(*) from dbo.te_cabecerarecibos_bk
group by cabrec_numrecibo
having count(*)>1


