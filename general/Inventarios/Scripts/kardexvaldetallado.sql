SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO





create procedure al_kardexvaldetallado_rpt
(
@base varchar(50),
@almacen varchar(2)
)
as
Declare @ncadena nvarchar(2000)
Declare @nparame nvarchar(2000)


Set @ncadena=N'SELECT a.alma,a.COD_ART,a.FEC_DOC,a.HOR_DOC,a.ING_SAL,a.COD_MOV,
		a.SER_LOT,a.TIP_TRANSA,a.NUM_DOC,a.CAN_ART,a.PRE_UNIT,
		a.COS_PRO,a.SAL_STOCK,b.adescri
		FROM ['+@base+'].dbo.al_kardex_val a
		INNER JOIN ['+@base+'].dbo.maeart b
		ON a.cod_art=b.ACODIGO
  	        Where a.ALMA=''' +@ALMACEN+''' and a.can_art > 0 '

--Set @nparame=N'@tipo varchar(2),@numero varchar(11)'
exec (@NCADENA)

--Execute sp_executesql @ncadena,@nparame,

--EXEC al_kardexvalorizadodetallado_rpt 'CAMTEX_TJ','01'



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

