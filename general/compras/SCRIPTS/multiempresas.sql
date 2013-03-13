
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO







--- drop proc al_listaproveedor_rep


CREATE  procedure al_listaproveedor_rep
@base as varchar(50)
as
declare @cadena as nvarchar(1000)
---declare @c as varchar(2)

---set @c='62'

set @cadena='Select A.PRVCCODIGO,A.PRVCNOMBRE,A.PRVCDIRECC,
              A.PRVCTELEF1,a.prvcruc 
             From ['+@base+'].dbo.MAEPROV A   
             '

execute(@cadena)










GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


'INSERT INTO ['+@Base+'].[dbo].[co_sistema] 
	 ( [sistemadescripcionempresa],
	 [sistemadescrcortaempresa],
	 [sistemaesttipodescrempresa],
	 [sistemadireccionempresa],
	 [sistemaempresaruc],
	 [monedacodigo],
	 [sistemactacomp],
	 [sistemaigv],
	 [usuariocodigo],
	 [fechaact],
	 [sistemasubasiento],
	 [sistemalibro],
	 [sistematipanal],
	 [sistemactatotal],
	 [sistemactaIGV], 
          sistemactaIES,  
          sistemactaRTA,
         sistematipoplan,
         sistemaoficina,
         permite_tc,
         sistemaactivaccostos,
         sistemaasientoenlinea,
         sistemamultiempresas ) 
 
  VALUES 
	(@sistemadescripcionempresa,
	 @sistemadescrcortaempresa,
	 @sistemaesttipodescrempresa,
	 @sistemadireccionempresa,
	 @sistemaempresaruc,
	 @monedacodigo,
	 @sistemactacomp,
	 @sistemaigv,
	 @usuariocodigo,
	 @fechaact,
	 @sistemasubasiento,
	 @sistemalibro,
	 @sistematipanal,
	 @sistemactatotal,
	 @sistemactaIGV,
         @sistemactaIES,
         @sistemactaRTA,
         @sistematipoplan,
         @sistemaoficina,
         @permite_tc,
         @sistemaactivaccostos,
         @sistemaasientoenlinea,
         @sistemamultiempresas) '
End
If @Op=2 
Begin 
Set @sqlcad=''+
'UPDATE ['+@Base+'].[dbo].[co_sistema] 
 SET  [sistemadescripcionempresa]	 = @sistemadescripcionempresa,
	 [sistemadescrcortaempresa]	 = @sistemadescrcortaempresa,
	 [sistemaesttipodescrempresa]	 = @sistemaesttipodescrempresa,
	 [sistemadireccionempresa]	 = @sistemadireccionempresa,
	 [sistemaempresaruc]	 = @sistemaempresaruc,
	 [monedacodigo]	 = @monedacodigo,
	 [sistemactacomp]	 = @sistemactacomp,
	 [sistemaigv]	 = @sistemaigv,
	 [usuariocodigo]	 = @usuariocodigo,
	 [fechaact]	 = @fechaact,
	 [sistemasubasiento]	 = @sistemasubasiento,
	 [sistemalibro]	 = @sistemalibro,
	 [sistematipanal]	 = @sistematipanal,
	 [sistemactatotal]	 = @sistemactatotal,
	 [sistemactaIGV]	 = @sistemactaIGV, 
          sistemactaIES          = @sistemactaIES,
          sistemactaRTA          = @sistemactaRTA,    
          sistematipoplan=@sistematipoplan,
          sistemaoficina=@sistemaoficina,
          permite_tc=@permite_tc,
          sistemaactivaccostos=@sistemaactivaccostos,
          sistemaasientoenlinea=@sistemaasientoenlinea,
          sistemamultiempresas=@sistemamultiempresas '
End
Exec sp_executesql @sqlcad,@sqlparm,
                           @sistemadescripcionempresa,
  			   @sistemadescrcortaempresa,
  			   @sistemaesttipodescrempresa,
  			   @sistemadireccionempresa,
  			   @sistemaempresaruc,
  			   @monedacodigo,
  			   @sistemactacomp 	,
  			   @sistemaigv,
  			   @usuariocodigo,
       			   @fechaact,
  			   @sistemasubasiento,
 	                   @sistemalibro,
	 		   @sistematipanal,
	 		   @sistemactatotal,
	 		   @sistemactaIGV,
                           @sistemactaIES,   
                           @sistemactaRTA,
                           @sistematipoplan,
                           @sistemaoficina, 
                           @permite_tc,
                           @sistemaactivaccostos,
                           @sistemaasientoenlinea,
                           @sistemamultiempresas


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

