if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[fn_xx]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[fn_xx]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[v_analiticoentidad]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[v_analiticoentidad]
GO

if exists (select * from dbo.systypes where name = N'fechaact')
exec sp_droptype N'fechaact'
GO

if exists (select * from dbo.systypes where name = N'numvalor')
exec sp_droptype N'numvalor'
GO

if exists (select * from dbo.systypes where name = N'usuariocodigo')
exec sp_droptype N'usuariocodigo'
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UW_ZeroDefault]') and OBJECTPROPERTY(id, N'IsDefault') = 1)
drop default [dbo].[UW_ZeroDefault]
GO

CREATE DEFAULT UW_ZeroDefault AS 0
GO
setuser
GO

EXEC sp_addtype N'fechaact', N'datetime', N'not null'
GO

setuser
GO

setuser
GO

EXEC sp_addtype N'numvalor', N'numeric(20,4)', N'null'
GO

setuser
GO

setuser
GO

EXEC sp_bindefault N'[dbo].[UW_ZeroDefault]', N'[numvalor]'
GO

setuser
GO

setuser
GO

EXEC sp_addtype N'usuariocodigo', N'nchar (8)', N'not null'
GO

setuser
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE FUNCTION [dbo].[fn_xx] ( @a float  )
RETURNS float
 WITH SCHEMABINDING
 AS  
BEGIN  
  return (@a * 1)	
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.v_analiticoentidad
AS
SELECT     a.analiticocodigo, a.tipoanaliticocodigo, b.*
FROM         dbo.ct_analitico a INNER JOIN
                      dbo.ct_entidad b ON a.entidadcodigo = b.entidadcodigo


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

