--dbo.ct_detcomprob2002
--dbo.ct_detcomprobxxxx

/*
 Objetivo: se crea a un campo fecha del documento de referencia
 Autor   : Fernando Cossio
 Fecha   : 17/10/2002	
*/
Alter Table ct_detcomprobxxxx Add detcomprobfecharef Datetime
Go
Alter Table ct_detcomprob2002 Add detcomprobfecharef Datetime
Go
ALTER     Proc ct_grabaautomatico_pro(
@base varchar(50),
@tabla varchar(50),
@mes  varchar(2), 
@comp varchar(10),
@asiento varchar(3) ,
@subasiento varchar(4))
as
DECLARE @sqlcad varchar(8000),@sqlcad2 varchar(8000), @sqlparm nvarchar(1000)
--Pararametros de la cadena
SET @sqlparm='@mes int,@comp varchar(10),@asiento varchar(3),
              @subasiento varchar(4)'

set @sqlcad='
	  SELECT 
       A.cabcomprobmes,A.cabcomprobnumero,A.subasientocodigo, 
       A.analiticocodigo,A.asientocodigo,detcomprobitem=''00000'',
       A.monedacodigo,A.centrocostocodigo,A.documentocodigo,
       A.operacioncodigo,
       cuentacodigo=cuentadistribucion,
       A.detcomprobnumdocumento,A.detcomprobfechaemision,A.detcomprobfechavencimiento,
       detcomprobglosa=''Asiento automatico'',
       debe=Round((isnull((case when A.detcomprobdebe > 0 then 
                 case when c.indicador=''D'' then (A.detcomprobdebe +  A.detcomprobhaber)    * (distribucionporcen/100) 
                 else 0 end
             else case when A.detcomprobhaber > 0 then 
                 	case when c.indicador=''H'' then (A.detcomprobdebe +  A.detcomprobhaber)    * (distribucionporcen/100) 
                 	else 0 end end
             end),0)),2),    
       haber=Round((isnull((case when A.detcomprobhaber > 0 then 
               case when c.indicador=''D'' then (A.detcomprobdebe +  A.detcomprobhaber) * (distribucionporcen/100) 
               else 0 end  
             else case when A.detcomprobdebe > 0 then  
                    case when c.indicador=''H'' then (A.detcomprobdebe +  A.detcomprobhaber) * (distribucionporcen/100) 
                    else 0 end end  
             end),0)),2),
       usshaber=Round((isnull((case when A.detcomprobhaber > 0 then 
                   case when c.indicador=''D'' then (A.detcomprobussdebe +  A.detcomprobusshaber) * (distribucionporcen/100) 
                   else 0 end  
                else case when A.detcomprobdebe > 0 then 
                        case when c.indicador=''H'' then (A.detcomprobussdebe +  A.detcomprobusshaber) * (distribucionporcen/100) 
                        else 0 end end 
                end),0)),2),
       ussdebe=Round((isnull((case when A.detcomprobdebe > 0 then 
                    case when c.indicador=''D'' then (A.detcomprobussdebe +  A.detcomprobusshaber) * (distribucionporcen/100) 
                    else 0 end 
                else case when A.detcomprobhaber > 0 then 
                    case when c.indicador=''H'' then (A.detcomprobussdebe +  A.detcomprobusshaber) * (distribucionporcen/100) 
                    else 0 end end
                end),0)),2),     
       A.detcomprobtipocambio,
       A.detcomprobruc,
       detcomprobauto=1,
       A.detcomprobformacambio,A.detcomprobajusteuser,A.plantillaasientoinafecto,
       A.detcomprobnlibro,A.detcomprobfecharef, 
      ID=IDENTITY(int, 1,1) 	
      INTO #tempo
FROM  ['+@base+'].dbo.ct_cuenta b,['+@base+'].dbo.['+@tabla+'] A,
      (select cuentacodigo,indicador=''D'',cuentadistribucion=distribucioncargo,distribucionporcen from ['+@base+'].dbo.ct_distribucion 
       where NOT (rtrim(distribucioncargo)='''' or 
       distribucioncargo is null)
        Union all
         select cuentacodigo,indicador=''H'',cuentadistribucion=distribucionabono,distribucionporcen from ['+@base+'].dbo.ct_distribucion 
         where NOT (rtrim(distribucionabono)='''' or 
         distribucionabono is null)) as C           
WHERE  B.cuentaestadodistribucion=1 and 
               B.cuentacodigo=A.cuentacodigo and
                B.cuentacodigo=C.cuentacodigo and  
                A.cabcomprobmes='+@mes+' and  
                A.cabcomprobnumero='''+@comp+''' and
                A.asientocodigo='''+@asiento+''' and
			    A.subasientocodigo='''+@subasiento+'''
       order by A.cabcomprobnumero,c.indicador 

'  
set @sqlcad2= '
if @@rowcount > 0 
Begin   
   declare @pend int
   select @pend=isnull(max(detcomprobitem),0) from ['+@base+'].dbo.['+@tabla+']
   where  cabcomprobmes='+@mes+' and  cabcomprobnumero='''+@comp+''' and  
            asientocodigo='''+@asiento+''' and  subasientocodigo='''+@subasiento+'''      

 insert into ['+@base+'].dbo.['+@tabla+']
   (  cabcomprobmes, cabcomprobnumero, subasientocodigo, 
      analiticocodigo, asientocodigo, 
      detcomprobitem,
      monedacodigo, centrocostocodigo, documentocodigo, operacioncodigo, cuentacodigo,
      detcomprobnumdocumento, detcomprobfechaemision, detcomprobfechavencimiento,
      detcomprobglosa, 
      detcomprobdebe, detcomprobhaber, detcomprobusshaber, detcomprobussdebe, detcomprobtipocambio, detcomprobruc,
      detcomprobauto, detcomprobformacambio, detcomprobajusteuser,plantillaasientoinafecto,detcomprobnlibro,detcomprobfecharef)
    select 
       cabcomprobmes,cabcomprobnumero,subasientocodigo, 
       analiticocodigo,asientocodigo,
       detcomprobitem=Replicate(''0'',5-len(@pend + id)) + rtrim(cast((@pend + id) as varchar(5))),
       monedacodigo,centrocostocodigo,documentocodigo, operacioncodigo,cuentacodigo,
       detcomprobnumdocumento,detcomprobfechaemision,detcomprobfechavencimiento,
       detcomprobglosa=''Asiento automatico'',
       debe,haber , usshaber, ussdebe, detcomprobtipocambio, detcomprobruc,
       detcomprobauto=1, detcomprobformacambio,detcomprobajusteuser,plantillaasientoinafecto,detcomprobnlibro,detcomprobfecharef
   from #tempo 
end'
exec(@sqlcad+@sqlcad2)
GO
ALTER    PROCEDURE ct_grabardetallecomprob_pro
	(@base varchar(30),
     @tabla varchar(30),
     @op int,
     @cabcomprobmes int	,
	 @cabcomprobnumero 	[char](10),
	 @subasientocodigo 	[char](4),
	@asientocodigo 	[char](3),
	@detcomprobitem 	[char](5)=null,
	 @analiticocodigo 	[char](15)=null, 	 
	 @monedacodigo 	[char](2)=null,
	 @centrocostocodigo 	[char](5)=null,
	 @documentocodigo 	[char](2)=null,
	 @operacioncodigo 	[char](2)=null,
	 @cuentacodigo 	[varchar](20)=null,
	 @detcomprobnumdocumento 	[Varchar](20)=null,
	 @detcomprobfechaemision 	[datetime]=null,
	 @detcomprobfechavencimiento 	[datetime]=null,
	 @detcomprobglosa 	[varchar](50)=null,
	 @detcomprobdebe 	[numvalor]=0,
	 @detcomprobhaber 	[numvalor]=0,
	 @detcomprobusshaber 	[numvalor]=0,
	 @detcomprobussdebe 	[numvalor]=0,
	 @detcomprobtipocambio 	[float]=0,
	 @detcomprobruc 	[char](13)=null,
	 @detcomprobauto 	[bit]=0,     
     @detcomprobformacambio char(2)=Null,  
     @detcomprobajusteuser  bit=0,
     @plantillaasientoinafecto bit =0,
     @tipdocref  char(2)=null, 
     @detcomprobnumref varchar(20)=null,  
     @detcomprobnlibro varchar(10)=Null,   
     @detcomprobfecharef datetime=Null

)

AS
DECLARE @sqlcad nvarchar(4000),@sqlparm nvarchar(1000)
--Pararametros de la cadena
SET @sqlparm='@cabcomprobmes int,@cabcomprobnumero char(10),@subasientocodigo char(4),'+
	   		 '@analiticocodigo char(15),@asientocodigo char(3),@detcomprobitem char(5),'+
	 		 '@monedacodigo char(2),@centrocostocodigo char(5),@documentocodigo char(2),'+
	 		 '@operacioncodigo char(2),@cuentacodigo varchar(20),@detcomprobnumdocumento Varchar(20),'+
	 		 '@detcomprobfechaemision datetime,@detcomprobfechavencimiento datetime,@detcomprobglosa varchar(50),'+
	 		 '@detcomprobdebe numvalor,@detcomprobhaber numvalor,@detcomprobusshaber numvalor,'+
	 		 '@detcomprobussdebe numvalor,@detcomprobtipocambio float,@detcomprobruc char(13),'+
	 		 '@detcomprobauto bit,@detcomprobformacambio char(2),
              @detcomprobajusteuser  bit,@plantillaasientoinafecto bit,
              @tipdocref char(2),@detcomprobnumref varchar(20),
              @detcomprobnlibro varchar(10),@detcomprobfecharef datetime'


IF @op=1 --Insertar Datos
BEGIN
	SET @sqlcad=''+
	'INSERT INTO '+'['+@base+'].[dbo].['+@tabla+'] 
   ([cabcomprobmes],
	 [cabcomprobnumero],
	 [subasientocodigo],
	 [analiticocodigo],
	 [asientocodigo],
	 [detcomprobitem],
	 [monedacodigo],
	 [centrocostocodigo],
	 [documentocodigo],
	 [operacioncodigo],
	 [cuentacodigo],
	 [detcomprobnumdocumento],
	 [detcomprobfechaemision],
	 [detcomprobfechavencimiento],
	 [detcomprobglosa],
	 [detcomprobdebe],
	 [detcomprobhaber],
	 [detcomprobusshaber],
	 [detcomprobussdebe],
	 [detcomprobtipocambio],
	 [detcomprobruc],
	 [detcomprobauto],
     [detcomprobformacambio],detcomprobajusteuser,
     plantillaasientoinafecto,tipdocref,detcomprobnumref,detcomprobnlibro,detcomprobfecharef)  
VALUES 
	(@cabcomprobmes,
	 @cabcomprobnumero,
	 @subasientocodigo,
	 @analiticocodigo,
	 @asientocodigo,
	 @detcomprobitem,
	 @monedacodigo,
	 @centrocostocodigo,
	 @documentocodigo,
	 @operacioncodigo,
	 @cuentacodigo,
	 @detcomprobnumdocumento,
	 @detcomprobfechaemision,
	 @detcomprobfechavencimiento,
	 @detcomprobglosa,
	 @detcomprobdebe,
	 @detcomprobhaber,
	 @detcomprobusshaber,
	 @detcomprobussdebe,
	 @detcomprobtipocambio,
	 @detcomprobruc,
	 @detcomprobauto,@detcomprobformacambio,
     @detcomprobajusteuser,@plantillaasientoinafecto,@tipdocref,@detcomprobnumref,@detcomprobnlibro,@detcomprobfecharef)'
END

IF @op=2 --Eliminar
BEGIN
	SET @sqlcad=''+
  	'DELETE FROM '+'['+@base+'].[dbo].['+@tabla+'] '+		
	'WHERE 
	( [cabcomprobmes]	 = @cabcomprobmes AND
	  [cabcomprobnumero]	 = @cabcomprobnumero AND
	  [subasientocodigo]	 = @subasientocodigo AND
	  [asientocodigo]	 = @asientocodigo )'
END
IF @op=3 --Recuperar los Datos
BEGIN
        SET @sqlcad=''+	
        'SELECT
	detcomprobitem,operacioncodigo,cuentacodigo,centrocostocodigo,
	B.tipoanaliticocodigo,A.analiticocodigo,detcomprobruc,documentocodigo,
	detcomprobnumdocumento,detcomprobfechaemision,
	detcomprobfechavencimiento,detcomprobglosa,monedacodigo,
	tcambio =detcomprobformacambio,valcambio=detcomprobtipocambio,
	indicador=case when detcomprobdebe > 0 then ''D'' else ''H'' end,
	montosol = isnull((case when detcomprobdebe > 0 then detcomprobdebe 
             			else detcomprobhaber end),0),
	montouss = isnull((case when detcomprobussdebe > 0 then detcomprobussdebe
             			 else detcomprobusshaber end),0),detcomprobauto,detcomprobajusteuser,plantillaasientoinafecto,tipdocref,detcomprobnumref,detcomprobnlibro,detcomprobfecharef

        FROM ['+@base+'].dbo.['+@tabla+']  A,['+@base+'].dbo.ct_analitico B
        WHERE
	A.cabcomprobmes=@cabcomprobmes and
	A.asientocodigo=@asientocodigo and
             A.subasientocodigo=@subasientocodigo and
	A.cabcomprobnumero=@cabcomprobnumero and
             A.analiticocodigo=B.analiticocodigo 
        order by A.cabcomprobmes,A.subasientocodigo,
             A.asientocodigo,A.cabcomprobnumero,A.detcomprobitem'	
END

Exec sp_executesql   @sqlcad,@sqlparm,
				   	 @cabcomprobmes,
					 @cabcomprobnumero,
					 @subasientocodigo,
					 @analiticocodigo,
					 @asientocodigo,
					 @detcomprobitem,
					 @monedacodigo,
					 @centrocostocodigo,
					 @documentocodigo,
					 @operacioncodigo,
					 @cuentacodigo,
					 @detcomprobnumdocumento,
					 @detcomprobfechaemision,
					 @detcomprobfechavencimiento,
					 @detcomprobglosa,
					 @detcomprobdebe,
					 @detcomprobhaber,
					 @detcomprobusshaber,
					 @detcomprobussdebe,
					 @detcomprobtipocambio,
					 @detcomprobruc,
					 @detcomprobauto,@detcomprobformacambio,
					 @detcomprobajusteuser,@plantillaasientoinafecto,@tipdocref,@detcomprobnumref,@detcomprobnlibro,
                     @detcomprobfecharef
GO

