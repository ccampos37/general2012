
MOV_ANALIS


select * from dbo.CT030101
where 
 MOV_REGIST+MOV_FILE+MOV_COMP in 


select a.*,yy.MOV_ANALIS 
--update dbo.CT030101 set MOV_ANALIS=yy.MOV_ANALIS
from dbo.CT030101 a,
	(select MOV_ANALIS,MOV_REGIST,MOV_FILE,MOV_COMP from 
	 	(select * from dbo.CT030101
		  where MOV_CUENTA like '42%') as ZZ ) as YY
where 
	a.MOV_REGIST=yy.MOV_REGIST and a.MOV_FILE=yy.MOV_FILE and a.MOV_COMP=yy.MOV_COMP
--order by MOV_REGIST,MOV_FILE,MOV_COMP


