Select B.NAME, * From SysColumns A ,Sysobjects B
where A.ID=B.ID and 
      A.xtype=60 and b.xtype='U' 
Order by B.Name 


where name like 'Valor%'
id 

      		Numeric Money  
      		Valor2  Valor
Xtype 		108      60         
xprec  	 	 20
xscale 	  	  2
xoffset      22      14  
colordet      3       2



select marfice.dbo.fn_datenumber(10,11,2002)

floor(cast(dsafasd as real))=marfice.dbo.fn_datenumber(10,11,2002)

Select id,name  From Sysobjects 
where xtype='U' and name like 'Xx%'


CREATE DEFAULT ceros AS 0
valor 5 48,2

update SysColumns 
Set xtype=108,
    xprec=20,
    xscale=2,
    --xoffset=22,
    xusertype=108,
    length=13,
    cdefault=731149650,    
    typestat=0            	
where ID=843150049
 and xtype=60



select * from SysColumns 
where ID=779149821


/*Inicio*/

CREATE DEFAULT ceros AS 0

Select id From Sysobjects 
Where name='ceros' 683149479


/*Actualizar Campos a Numerico */
update SysColumns 
Set --xtype=108,
    --xprec=20,
    --xscale=2,    
    --xusertype=108,
    --length=13,
    cdefault=683149479
    --typestat=0            	
From SysColumns A ,Sysobjects B
where A.ID=B.ID and 
      A.xtype=108 and b.xtype='U' 


/*Fin*/




Declare @xtabla varchar(100),@xcampo varchar(100),@sqlcad varchar(500)     

Declare Tabla Cursor for
Select Tabla=B.NAME,Campo=A.Name From SysColumns A ,Sysobjects B
where A.ID=B.ID and 
      A.xtype=60 and b.xtype='U' 
Order by B.Name 

open tabla
fetch next from tabla into @xtabla,@xcampo
while @@fetch_status=0
Begin
	print 'Actualizando Tabla :'+@xtabla+' Campo:'+@xcampo  
	set @sqlcad='Alter Table ['+@xtabla+'] Alter Column ['+@xcampo+'] Numeric(20,2) ' 
    exec(@sqlcad)
	fetch next from tabla into @xtabla,@xcampo
End 
Close tabla
Deallocate tabla




##CALCINPUT" & Trim(VGL_COMPUTER)
select * from ##CALCINPUTDESARROLLO4