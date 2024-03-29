USE [marfice]
GO
/****** Object:  StoredProcedure [dbo].[ct_LibroCajaBancos_rpt]    Script Date: 10/28/2011 10:00:42 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
/*
drop  proc ct_libroCajaBancos_rpt
exec ct_libroCajaBancos_rpt 'planta_casma','03','2011','01','00','FORMATO 01.02'
*/

ALTER proc [dbo].[ct_LibroCajaBancos_rpt]
( 
    @base   varchar(50),
    @Empresa varchar(2),
    @anno   varchar(4),
    @mesact varchar(2),
    @mesant varchar(2),
    @formato varchar(20),
    @ctagastos varchar(10)='676900',
    @ctaingresos varchar(10)='776100'              
)
as
declare @sqlcad varchar(5000)
declare @sqlcad1 varchar(5000)
declare @sqlcad2 varchar(5000)
declare @cad1 varchar(100)
declare @cad2 varchar(100)
declare @cad3 varchar(100)

Set @cad3=(select formatocuerntacomodin from dbo.ct_formatos where formatocodigo=@formato )  
---execute(@cad3)
set @sqlcad='declare @cad1 varchar(100)
declare @cad2 varchar(100)

Set @cad1=(select formatodescripcion1 from '+@base+'.dbo.ct_formatos where formatocodigo='''+@FORMATO+''' )
Set @cad2=(select formatodescripcion2 from '+@base+'.dbo.ct_formatos where formatocodigo='''+@FORMATO+''' )'

IF cast(@mesant as integer)>0 
BEGIN
set @sqlcad=@sqlcad+' select zz.* , Descripcionasociada=bb.cuentadescripcion, descripcionmes=dbo.fn_DescripcionMes('+ @mesact+' ),
     aaaa='''+@anno+''' from 
(
SELECT a.empresacodigo,formatodescripcion1=@cad1,formatodescripcion2=@cad2,A.detcomprobfechaemision,A.cabcomprobnumero,Aa.documentocodigo,
detcomprobnumdocumento= dbo.fn_FormatoNumDoc(Aa.detcomprobnumdocumento),
       A.tipdocref,A.detcomprobnumref,detcomprobglosa=case when aa.analiticocodigo=''00'' then aa.detcomprobglosa else ee.entidadrazonsocial end ,
       A.detcomprobtipocambio,Aa.detcomprobusshaber-Aa.detcomprobussdebe as ComprobUSS,
       detcomprobdebe=case when aa.detcomprobauto=''0''  then Aa.detcomprobhaber else Aa.detcomprobdebe  end ,
       detcomprobhaber=case when aa.detcomprobauto=''0''  then Aa.detcomprobdebe else Aa.detcomprobhaber  end ,
d.formadepagocodigo,formadepagodescripcion,moneda=aa.monedacodigo, --case when b.tipoajuste=''01'' then ''02'' else ''01'' end,
       b.tipoajuste,A.detcomprobussdebe,A.detcomprobusshaber,SaldoDebe=C.saldoacumdebe' +@mesant+ ',SaldoHaber=C.saldoacumhaber' +@mesant+ ',
       SaldoIni=(C.saldoacumdebe' +@mesant+ '- C.saldoacumhaber' +@mesant+ '),SaldoUS =(C.saldoacumussdebe' +@mesant+ '- C.saldoacumusshaber' +@mesant+ '),'
END
ELSE
BEGIN
set @sqlcad=@sqlcad+' select zz.* , Descripcionasociada=bb.cuentadescripcion , descripcionmes=dbo.fn_DescripcionMes('+ @mesact+' ),
        aaaa='''+@anno+'''  from 
(
SELECT a.empresacodigo,formatodescripcion1=@cad1,formatodescripcion2=@cad2,A.detcomprobfechaemision,A.cabcomprobnumero,Aa.documentocodigo,
detcomprobnumdocumento= dbo.fn_FormatoNumDoc(Aa.detcomprobnumdocumento),
       A.tipdocref,A.detcomprobnumref,detcomprobglosa=case when aa.analiticocodigo=''00'' then aa.detcomprobglosa else ee.entidadrazonsocial end ,
       A.detcomprobtipocambio,Aa.detcomprobusshaber-Aa.detcomprobussdebe as ComprobUSS,
       detcomprobdebe=case when aa.detcomprobauto=''0''  then Aa.detcomprobhaber else Aa.detcomprobdebe  end ,
       detcomprobhaber=case when aa.detcomprobauto=''0''  then Aa.detcomprobdebe else Aa.detcomprobhaber  end ,
       d.formadepagocodigo,formadepagodescripcion,moneda=aa.monedacodigo,--case when b.tipoajuste=''01'' then ''02'' else ''01'' end,
       b.tipoajuste,A.detcomprobussdebe,A.detcomprobusshaber,       SaldoDebe=C.saldoacumdebe' +@mesant+ ',SaldoHaber=C.saldoacumhaber' +@mesant+ ',
       SaldoIni=(C.saldoacumdebe' +@mesant+ '- C.saldoacumhaber' +@mesant+ '),SaldoUS =(C.saldoacumussdebe' +@mesant+ '- C.saldoacumusshaber' +@mesant+ '),'
END

set @sqlcad1='SaldoIniD=(C.saldoussdebe' +@mesant+ '- C.saldousshaber' +@mesant+ '),
    SaldoFin=Aa.detcomprobdebe-Aa.detcomprobhaber,A.cuentacodigo,B.cuentadescripcion,A.monedacodigo,Cuenta2=left(A.cuentacodigo,2),
    Cuentaasociada=case when aa.detcomprobauto=''1'' then 
                        case when ( A.detcomprobdebe - A.detcomprobhaber )> 0  then '''+@ctaingresos+''' else '''+@ctagastos+'''    end
                    else aa.cuentacodigo end ,
                     cc.cabcomprobfeccontable
    FROM  [' +@base+ '].dbo.[ct_detcomprob' +@anno+ '] A 
    INNER JOIN [' +@base+ '].dbo.[ct_cuenta] B ON a.empresacodigo=b.empresacodigo and A.cuentacodigo = B.cuentacodigo 
    INNER JOIN [' +@base+ '].dbo.[ct_saldos' + @anno+ '] C ON a.empresacodigo=c.empresacodigo and A.cuentacodigo = C.cuentacodigo
    Left JOIN [' +@base+ '].dbo.[gr_documento] d ON isnull(A.documentocodigo,''003'') = d.documentocodigo
    Left JOIN [' +@base+ '].dbo.[ct_formadepago] e ON d.formadepagocodigo=e.formadepagocodigo
    Inner join [' +@base+ '].dbo.[ct_detcomprob' +@anno+ '] aa 
         on a.empresacodigo+a.asientocodigo+a.subasientocodigo+a.cabcomprobnumero=aa.empresacodigo+aa.asientocodigo+aa.subasientocodigo+aa.cabcomprobnumero          
    INNER JOIN [' +@base+ '].dbo.[ct_cabcomprob'+@anno+'] cc 
         on a.empresacodigo+a.asientocodigo+a.subasientocodigo+a.cabcomprobnumero=cc.empresacodigo+cc.asientocodigo+cc.subasientocodigo+cc.cabcomprobnumero   
    LEFT JOIN [' +@base+ '].dbo.[ct_analitico] dd on aa.analiticocodigo=dd.analiticocodigo
    left join [' +@base+ '].dbo.[ct_entidad] ee on dd.entidadcodigo=ee.entidadcodigo
WHERE A.empresacodigo='''+@empresa+'''  AND A.cabcomprobmes=''' +@mesact+ ''' and left(aa.cuentacodigo,2) <''79'' 
     and left(a.cuentacodigo,3) in (' + @cad3 + ')  and 
         (a.detcomprobitem <> aa.detcomprobitem and aa.detcomprobauto=''0'' or 
         a.detcomprobitem = aa.detcomprobitem and aa.detcomprobauto=''1''  ) 
   ) zz
 INNER JOIN [' +@base+ '].dbo.[ct_cuenta] Bb ON zz.empresacodigo=bb.empresacodigo and zz.Cuentaasociada = Bb.cuentacodigo   
	ORDER BY zz.cuentacodigo,zz.cabcomprobfeccontable,zz.cabcomprobnumero,zz.cuentaasociada '
	 execute (@sqlcad +@sqlcad1)
---     exec ct_libroCajaBancos_rpt 'conta2010','02','2010','12','11','FORMATO 01.01'














