select cabcomprobnumero,count(*) from ct_cabcomprob2003 
where cabcomprobmes=2
group by  cabcomprobnumero
having  count(*)>1
order by 1

--select * from ct_detcomprob2003
where cabcomprobnumero='0407200801'
--0407200103 
--0407200104

select distinct(cabcomprobnumero) from ct_detcomprob2003
where cabcomprobmes=4 and 
	detcomprobnumdocumento like '003%' and 
	subasientocodigo='0004' and asientocodigo like '07%'


select max(cabcomprobnumero) from ct_detcomprob2003
where cabcomprobnumero like '04071%'
where cabcomprobnumero='0407200801'
