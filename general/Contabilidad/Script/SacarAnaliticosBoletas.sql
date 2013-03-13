--select * from ct_detcomprob2003 

select ZZ.*,a.entidadrazonsocial 
	from
		dbo.v_analiticoentidad a,
		(select Num=count(*),analiticocodigo from ct_detcomprob2003 
			where asientocodigo like '07%' and cabcomprobmes=1 and
   			detcomprobnumdocumento like '008%' and documentocodigo in ('03')
			group by analiticocodigo) as ZZ
where ZZ.analiticocodigo=a.analiticocodigo

select ZZ.*,a.entidadrazonsocial 
	from
		dbo.v_analiticoentidad a,
		(select Num=count(*),analiticocodigo from ct_detcomprob2003 
			where asientocodigo like '07%' and cabcomprobmes=1 and
			   detcomprobnumdocumento like '002%' and documentocodigo in ('03')
			group by analiticocodigo) as ZZ
where ZZ.analiticocodigo=a.analiticocodigo

select ZZ.*,a.entidadrazonsocial 
	from
		dbo.v_analiticoentidad a,
		(select Num=count(*),analiticocodigo from ct_detcomprob2003 
			where asientocodigo like '07%' and cabcomprobmes=1 and
			   detcomprobnumdocumento like '003%' and documentocodigo in ('03')
			group by analiticocodigo ) as ZZ
where ZZ.analiticocodigo=a.analiticocodigo

