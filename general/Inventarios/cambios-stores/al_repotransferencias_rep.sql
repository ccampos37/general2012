alter procedure al_repotransferencias_rep
@base varchar(50),
@fini varchar(10),
@ffin varchar(10)
as
declare @ncadena varchar(1000)
declare @tr varchar(2)

set @tr='TR'

set @ncadena='
	select carftdoc,carfndoc,caalma,catd,canumdoc,catipmov,cafecdoc,
	       decantid,decanref1,decodigo,adescri 
	from ['+@base+'].dbo.movalmcab inner join ['+@base+'].dbo.movalmdet
	on caalma=dealma and catd=detd and canumdoc=denumdoc
	inner join ['+@base+'].dbo.maeart
	on decodigo=acodigo
	where cafecdoc>='''+@fini+''' and cafecdoc<='''+@ffin+''' and 
	carftdoc='''+@tr+'''
	order by carftdoc,carfndoc,caalma,catipmov'
execute(@ncadena)



execute al_repotransferencias_rep
@base='bdcomunbazo',
@fini='01/01/2003',
@ffin='29/01/2003'


select movalmcab.carftdoc,movalmcab.carfndoc,movalmcab.caalma,movalmcab.catd,movalmcab.canumdoc,movalmcab.catipmov,
	 movalmdet.decantid,movalmdet.decanref1,movalmdet.decodigo,maeart.adescri 
	from [bdcomunbazo].dbo.movalmcab movalmcab
        inner join [bdcomunbazo].dbo.movalmdet movalmdet
	on caalma=dealma and catd=detd and canumdoc=denumdoc
	inner join [bdcomunbazo].dbo.maeart maeart
	on decodigo=acodigo
	where movalmcab.cafecdoc>='01/01/2003' and movalmcab.cafecdoc<='29/01/2003' and 
	movalmcab.carftdoc='TR'
	order by movalmcab.carftdoc,movalmcab.carfndoc
