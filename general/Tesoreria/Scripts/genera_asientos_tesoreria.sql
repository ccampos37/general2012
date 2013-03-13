SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO









ALTER            proc te_GeneraAsientosTesoreria_pro 
--Declare
  	@Baseventa    	varchar(100),
  	@Baseconta 		varchar(100),
  	@Asiento	    	varchar(15), 
  	@SubAsiento 	varchar(15),
  	@Libro   		varchar(2),         
  	@Mes     		varchar(2),
  	@Ano     		varchar(4),        
  	@tipanal      	varchar(3), 
  	@Compu   		varchar(50),
  	@Usuario 		varchar(20),
	@TipoMov			varchar(1)	
as    
/*
  Set @Baseventa='VENTAS_PRUEBA' 
  Set @Baseconta='CONTAPRUEBA'
  Set @Asiento='012'
  Set @SubAsiento='0001'
  SET @Libro='01'
  Set @Mes='12'
  Set @Ano='2002' 
  set @tipanal='001'
  Set @Compu='Desarrollo3'
  Set @Usuario='TESOR01' 
*/

exec marfice_ventas.dbo.te_GeneraTempContable @Baseventa,@Baseconta,'776101','976101','Desarrollo3',@mes,@ano,@TipoMov
--exec marfice_ventas.dbo.te_GeneraTempContable 'camtex_tinto','Contaprueba','776101','976101','Desa3','01','2003','E'


Declare @SqlCad varchar(8000),@SqlCad2 varchar(8000)

Set @SqlCad=
'If Exists(Select name from tempdb..sysobjects where name=''##tmpgenasientocab'+@compu+''') 
    Drop Table [##tmpgenasientocab'+@compu+'] 

 Select numprovi=A.cabrec_numrecibo,
        cabcomprobmes='+@Mes+',
        cabcomprobnumero=cast(''0'' as bigint) ,correlibro=IDENTITY(bigint,1,1),
        cabcomprobfeccontable=cast (''31/01/2003'' as datetime) ,
        subasientocodigo='''+@SubAsiento+''',usuariocodigo='''+@Usuario+''',estcomprobcodigo=''01'',
        asientocodigo='''+@Asiento+'''   ,cabcomprobobservaciones='' '',
        fechaact=getdate(),cabcomprobglosa=''Recibo de Egreso''+A.cabrec_numrecibo,
        cabcomprobtotdebe=0,
        cabcomprobtothaber=0,
        cabcomprobtotussdebe=0,
        cabcomprobtotusshaber=0,cabcomprobgrabada=0,cabcomprobnref='' '',
        cabcomprobnlibro=0 Into [##tmpgenasientocab'+@Compu+'] 
 From   [##tmpAsientosConta'+@Compu+'] A
 Group by A.cabrec_numrecibo'
Exec(@SqlCad) 

/*
Declare @Cad varchar(1000)
set @cad='
update ##tmpgenasientocab' +@Compu+ '
  set cabcomprobfeccontable=(select top 1 FechaCancela 
		from ##tmpCuentaCajaDesarrollo3 A where A.cabrec_numrecibo=B.cabrec_numrecibo)
  from ##tmpAsientosConta' +@Compu+ ' B '

exec(@cad)
*/

--select * from ##tmpCuentaCajaDesarrollo3

Set @SqlCad='
 --Seleccion del Detalle
 If Exists(Select name from tempdb..sysobjects where name=''##tmpgenasientodet'+@compu+''') 
    Drop Table [##tmpgenasientodet'+@compu+']  
  
 Select *,numfila=IDENTITY(BigInt,1,1) Into [##tmpgenasientodet'+@Compu+'] From 
 (select numprovi=A.cabrec_numrecibo,
       cabcomprobmes='+@Mes+',
       cabcomprobnumero='' '',subasientocodigo='''+@SubAsiento+''',
       analiticocodigo=A.clientecodigo,                      
       asientocodigo='''+@Asiento+''',detcomprobitem=Replicate(''0'',5-len(A.detrec_item))+rtrim(ltrim(cast(A.detrec_item as varchar(10)))),monedacodigo=A.monedacodigo,
       centrocostocodigo=''00'',
       documentocodigo=A.detrec_tipodoc_concepto,operacioncodigo=''03'',cuentacodigo=A.cuenta,
       detcomprobnumdocumento=A.detrec_numdocumento,detcomprobfechaemision=A.FechaEmision,
       detcomprobfechavencimiento=A.FechaCancela,detcomprobglosa=A.detrec_observacion,            
       detcomprobdebe=isnull(A.DebeS ,0),
       detcomprobhaber=isnull(A.HaberS,0),
       detcomprobusshaber= isnull(DebeD,0), 
       detcomprobussdebe=isnull(HaberD,0),  
       detcomprobtipocambio=A.TipoCambio , detcomprobruc=space(11),
       detcomprobauto=0, detcomprobformacambio=''02'', 
       detcomprobajusteuser=0, plantillaasientoinafecto=0,
       tipdocref=space(2), detcomprobnumref=space(11),
       detcomprobconci=0, detcomprobnlibro='' '' ,
       detcomprobfecharef=null,cabprovinconta=A.cabrec_numrecibo
from  [##tmpAsientosConta'+@compu+']  A
     ) as XX '     
exec(@SqlCad) 

--exec marfice.dbo.cc_insertanalitico_pro @Baseconta,@Baseventa,@Compu  
exec marfice.dbo.te_InsertaAnaliticoEgreso_pro @Baseconta,@Baseventa,@Compu,'001'
exec marfice.dbo.te_InsertaAnaliticoIngreso_pro @Baseconta,@Baseventa,@compu,'002'


Declare @CtaReg BigInt
Exec('
Declare CuentaReg Cursor for 
Select CtaReg=Count(*) From [##tmpgenasientocab'+@compu+']')

Open CuentaReg
Fetch Next from CuentaReg into @CtaReg
Close CuentaReg
Deallocate CuentaReg

If @CtaReg=0 
Begin
   Print 'No Existen Registros para generar a contabilidad '	
--   Return 0
End
--El correlativo es por libros 
--se tiene generar un temporal por cada asiento y su correlativo
--para cada correlativo de cada libro

--collate  Modern_Spanish_CI_AI
SET @SqlCad='
 If Exists(Select name from tempdb..sysobjects where name=''##tmpcorrela'+@compu+''') 
    Drop Table [##tmpcorrela'+@compu+']  
 
 select MaxAsi=asientonumcorr'+@MES+',A.Asiento2,Ultimo=asientonumcorr'+@MES+' 
 Into [##tmpcorrela'+@compu+']
 from ['+@BaseConta+'].dbo.te_pasientocab A,['+@BaseConta+'].dbo.ct_asientocorre B
 where 
	  B.asientoanno='''+@Ano+''' and 	
      A.Asiento2  =  B.asientocodigo  ' 
EXEC(@SqlCad)


SET @SqlCad='  
Declare @Asiento varchar(3),@Numprovi VARCHAR(20)
Declare Correla cursor for 
select asientocodigo,numprovi from [##tmpgenasientocab'+@compu+']
order by asientocodigo

Open Correla
fetch next from Correla into @Asiento,@Numprovi

While @@Fetch_Status=0 
Begin 
   update [##tmpgenasientocab'+@compu+']
   set cabcomprobnumero=isnull(B.Ultimo,0) +1
   From  [##tmpgenasientocab'+@compu+'] A,
         [##tmpcorrela'+@compu+'] B
   Where A.Asientocodigo collate  Modern_Spanish_CI_AI  =B.Asiento2 collate  Modern_Spanish_CI_AI  and 
         numprovi=@Numprovi
   
   Update  [##tmpcorrela'+@compu+'] 
   Set  Ultimo=ISNULL(Ultimo,0)+1 
   Where  Asiento2=@Asiento
   fetch next from Correla into @Asiento,@Numprovi		
End
Close Correla
Deallocate Correla '
EXEC(@SqlCad)     

/*
declare @cadsql varchar(2000)
--set @NombrePC='Desarrollo3'
set @cadsql='update [##tmpgenasientodet'+@compu+ '] set analiticocodigo=''00'' 
 	 where rtrim(ltrim(isnull(analiticocodigo,'''')))='''''
exec(@cadsql)
*/


Set @SqlCad2=' '+ 
'Update [##tmpgenasientodet'+@compu+']
 Set
     detcomprobtipocambio=tipocambioventa,
     detcomprobformacambio=''02'',
     detcomprobdebe=isnull(case when A.monedacodigo =''01''     then  A.detcomprobdebe else Case when left(A.cuentacodigo,2) in (''77'',''97'') then 0 else round(A.detcomprobussdebe*tipocambioventa,2) end end  ,0)   ,
     detcomprobussdebe=isnull(case when A.monedacodigo =''02''  then  A.detcomprobussdebe else Case when left(A.cuentacodigo,2) in (''77'',''97'') then  0 else round(A.detcomprobdebe/tipocambioventa,2) end end ,0) , 
     detcomprobhaber=isnull(case when A.monedacodigo =''01''    then  A.detcomprobhaber else Case when left(A.cuentacodigo,2) in (''77'',''97'') then  0 else round(A.detcomprobusshaber*tipocambioventa,2) end end,0), 
     detcomprobusshaber=isnull(case when A.monedacodigo =''02'' then  A.detcomprobusshaber else Case when left(A.cuentacodigo,2) in (''77'',''97'') then 0 else round(A.detcomprobhaber/tipocambioventa,2) end end,0)  
 From [##tmpgenasientodet'+@compu+'] A,
              ['+@baseconta+'].dbo.gr_documento B,
              ['+@baseconta+'].dbo.ct_tipocambio C

 Where A.documentocodigo  =B.documentocodigo  and  
      (Case When B.documentonotacredito=0 then A.detcomprobfechaemision 
       Else A.detcomprobfecharef end) =C.tipocambiofecha ' 


Set @SqlCad2='
 Declare @MaxLibro Bigint
 Select @MaxLibro=libronumcorr'+@mes+' from '+@BaseConta+'.dbo.ct_librocorre 
 where  librocodigo='''+@Libro+''' and libroanno='''+@Ano+''' 

 Insert Into ['+@BaseConta+'].dbo.ct_cabcomprob'+@Ano+'
(cabcomprobmes, cabcomprobnumero , cabcomprobfeccontable, subasientocodigo, 
 usuariocodigo, estcomprobcodigo, asientocodigo, cabcomprobobservaciones, 
 fechaact, cabcomprobglosa, cabcomprobtotdebe, cabcomprobtothaber,
 cabcomprobtotussdebe, cabcomprobtotusshaber, cabcomprobgrabada,
 cabcomprobnref, cabcomprobnlibro,cabcomprobnprovi)

 Select  cabcomprobmes,
       comprobnumero='''+@mes+'''+asientocodigo+replicate(''0'',5-len(cabcomprobnumero))+ltrim(rtrim(cast(cabcomprobnumero as varchar(20))))     
       , cabcomprobfeccontable, subasientocodigo, 
        usuariocodigo, estcomprobcodigo, asientocodigo, cabcomprobobservaciones, fechaact, 
        cabcomprobglosa, cabcomprobtotdebe, cabcomprobtothaber, cabcomprobtotussdebe, 
        cabcomprobtotusshaber, cabcomprobgrabada, cabcomprobnref,
        comprobnlibro='''+@mes+'''+'''+@libro+'''+replicate(''0'',6-len(@MaxLibro+correlibro))+ltrim(rtrim(cast(@MaxLibro+correlibro as varchar(20)))),
        ''TES''+numprovi  
 from [##tmpgenasientocab'+@compu+'] A 


Insert Into ['+@BaseConta+'].dbo.ct_detcomprob'+@Ano+'
(cabcomprobmes, cabcomprobnumero, subasientocodigo, analiticocodigo, asientocodigo,
 detcomprobitem, monedacodigo, centrocostocodigo, documentocodigo, operacioncodigo,
 cuentacodigo, detcomprobnumdocumento, detcomprobfechaemision, detcomprobfechavencimiento,
 detcomprobglosa, detcomprobdebe, detcomprobhaber, detcomprobusshaber, detcomprobussdebe,
 detcomprobtipocambio, detcomprobruc, detcomprobauto, detcomprobformacambio,
 detcomprobajusteuser, plantillaasientoinafecto, tipdocref,
 detcomprobnumref, detcomprobconci, detcomprobnlibro, detcomprobfecharef)

Select Distinct
       A.cabcomprobmes, 
       comprobnumero='''+@mes+'''+A.asientocodigo+replicate(''0'',5-len(B.cabcomprobnumero))+ltrim(rtrim(cast(B.cabcomprobnumero as varchar(10)))),
       A.subasientocodigo,
       analitico=case when left(A.cuentacodigo,2)=''42'' then 
                         case Rtrim(ltrim(isnull(cli.clienteruc,''''))) 
                         when ''00000000000'' then cli.clientecodigo 
                         When '''' Then cli.clientecodigo
                         else  cli.clienteruc end + ''001''
                   Else ''00'' end,
       A.asientocodigo,
       A.detcomprobitem, A.monedacodigo, A.centrocostocodigo, A.documentocodigo, A.operacioncodigo,
       A.cuentacodigo, A.detcomprobnumdocumento, A.detcomprobfechaemision, A.detcomprobfechavencimiento,
       detcomprobglosa=left(A.detcomprobglosa,50), A.detcomprobdebe, A.detcomprobhaber, A.detcomprobusshaber, A.detcomprobussdebe,
       A.detcomprobtipocambio,
		 detcomprobruc=isnull(cli.clienteruc,'''')  , A.detcomprobauto, A.detcomprobformacambio,
       A.detcomprobajusteuser, A.plantillaasientoinafecto,
       tipdocref= case when rtrim(isnull(A.tipdocref,''00''))='''' then ''00'' else isnull(A.tipdocref,''00'') end,
       A.detcomprobnumref, A.detcomprobconci, 
       comprobnlibro='''+@mes+'''+'''+@libro+'''+replicate(''0'',5-len(@MaxLibro+B.correlibro))+ltrim(rtrim(cast(@MaxLibro+B.correlibro as varchar(10))))
       , A.detcomprobfecharef
from [##tmpgenasientodet'+@compu+'] A, 
     [##tmpgenasientocab'+@compu+'] B,['+@Baseventa+'].dbo.cp_proveedor cli      
Where ltrim(rtrim(A.numprovi))=rtrim(ltrim(B.numprovi)) 
      and  rtrim(ltrim(A.analiticocodigo))=rtrim(ltrim(Cli.clientecodigo)) '
     

print (@SqlCad2)


/*
--Generar Automaticos 
--Generando Asientos Automaticos y Calculando el total del comprobante
Declare @Xcabcomprobnumero varchar(10),@Xasientocodigo varchar(3),
        @Xsubasientocodigo varchar(4),@Xtabla varchar(50)


Set @Xtabla='ct_detcomprob'+@Ano
set @Sqlcad='Declare GenAuto Cursor for 
select B.cabcomprobnumero,B.asientocodigo,B.subasientocodigo 
from [##tmpgenasientocab'+@compu+'] A,['+@BaseConta+'].dbo.ct_cabcomprob'+@Ano+' B
Where
 ''TES''+rtrim(ltrim(A.numprovi))=rtrim(ltrim(B.cabcomprobnprovi))' 
Exec(@Sqlcad)
Open GenAuto

Fetch next from GenAuto into @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo

While @@Fetch_status=0 
Begin
    Exec marfice.dbo.ct_grabaautomatico_pro @baseconta,@Xtabla,@mes,
    @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo                                   

    Fetch next from GenAuto into @Xcabcomprobnumero,@Xasientocodigo,@Xsubasientocodigo	
End
Close GenAuto
Deallocate GenAuto
*/


--Actualizar Correlativos de Asientos 

Set @SqlCad=''+
'Update ['+@BaseConta+'].dbo.ct_asientocorre 
 Set asientonumcorr'+@Mes+'= B.Ultimo
 From  
 ['+@BaseConta+'].dbo.ct_asientocorre A,
 [##tmpcorrela'+@Compu+'] B
 Where A.asientocodigo  =B.Asiento2  and 
       A.asientoanno='''+@Ano+'''  
  '  
Exec(@SqlCad) 
--Actualiza correlativo de Libros
Set @SqlCad=''+
'Update  ['+@BaseConta+'].dbo.ct_librocorre
 Set libronumcorr'+@Mes+'=libronumcorr'+@Mes+'+ 
           (Select count(*)  from ##tmpgenasientocab'+@COMPU+') 
 Where librocodigo='''+@Libro+''' and  libroanno='''+@Ano+''''
Exec(@SqlCad)

Exec marfice.dbo.cc_actualizacab @Baseconta,@Ano,@Mes,@Asiento


/*Ejecutar el Store de Generación de Asientos para Contabilidad*/
--EXEC te_GeneraAsientosTesoreria_pro 'camtex_tinto','contaprueba','012','0001','01','01','2003','001','Desarrollo3','TESOR01','E'


/*
select * From contaprueba..ct_cabcomprob2003 
  Where cabcomprobmes=1 and asientocodigo='012' 

delete from contaprueba..ct_cabcomprob2003 
  Where cabcomprobmes=1 and asientocodigo='012' 

select * From contaprueba..ct_detcomprob2003 
  Where cabcomprobmes=1 and asientocodigo='012' 

select sum(detcomprobdebe),sum(detcomprobhaber) From contaprueba..ct_detcomprob2003 
  Where cabcomprobmes=1 and asientocodigo='012' 

*/














GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

