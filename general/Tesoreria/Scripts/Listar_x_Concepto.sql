select a.*,b.* from te_detallerecibos b, te_cabecerarecibos a
where b.detrec_tipodoc_concepto='12' and 
		month(b.detrec_fechacancela)=12 and year(b.detrec_fechacancela)=2002 and
		a.cabrec_numrecibo=b.cabrec_numrecibo and
		a.cabrec_estadoreg<>'1' and b.detrec_monedadocumento=''
order by 1
