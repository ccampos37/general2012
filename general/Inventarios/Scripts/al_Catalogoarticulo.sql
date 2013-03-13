SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



ALTER       procedure al_listaarticulo_rep
@base as varchar(50),
@tipo as varchar(1),
@filtro as varchar(150)

as
declare @cadena as nvarchar(2000)
declare @cadena1 as nvarchar(500)

if @tipo='1' or @tipo='3'
   Begin
        SET @cadena1 =' Order by Acodigo'
   End
if @tipo='2' 
   Begin
        SET @cadena1 =' Order by Acodigo'
   End 

SET @cadena =N'Select acodigo,adescri,acodigo2,adescri2,afamilia
             From ['+@base+'].dbo.MAEART A Where '+@filtro+' '+@cadena1+''

execute(@cadena)

-- EXEC al_listaarticulo_rep 'CAMTEX_TJ','1','ACODIGO >=**'











GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

