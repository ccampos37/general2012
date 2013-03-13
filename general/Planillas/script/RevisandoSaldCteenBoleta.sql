SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

--ALTER  PROCEDURE [ASISTMP3] 

--ALTER  proc ArmaBol
Declare
@Base Varchar(100)

Set @Base='Camtex'

--AS 
DECLARE @CREAR VARCHAR(700)
SET @CREAR ='
            SELECT *              
	        FROM RPTBOLETAS TABLA1, 
           		 ['+@base+'].dbo.TRABAJADORES TABLA2,
		           (select A.Codtrab,B.CODIGO,B.MES,B.FECHAINI,B.FECHAFIN,B.FECHAPAGO,
				    	   Dias=Count(*)  
					from   ['+@base+'].dbo.asis2000 A,['+@base+'].dbo.NOMBOL B
					Where  A.Dia between B.FECHAINI and B.FECHAFIN and  
					       concepto=''HORASTRB'' and Valor > 0
					Group By A.Codtrab,B.CODIGO,B.MES,B.FECHAINI,B.FECHAFIN,B.FECHAPAGO)AS B            
	      WHERE TABLA1.CODTRAB=TABLA2.CODTRAB AND 
	            TABLA1.CODTRAB*=B.CODTRAB and 
	            TABLA1.NOMBOL*=B.Codigo '
PRINT (@CREAR)
--EXECUTE (@CREAR)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



SELECT *              
	        FROM RPTBOLETAS TABLA1, 
           		 [Camtex].dbo.TRABAJADORES TABLA2,
		           (select A.Codtrab,B.CODIGO,B.MES,B.FECHAINI,B.FECHAFIN,B.FECHAPAGO,
				    	   Dias=Count(*)  
					from   [Camtex].dbo.asis2000 A,[Camtex].dbo.NOMBOL B
					Where  A.Dia between B.FECHAINI and B.FECHAFIN and  
					       concepto='HORASTRB' and Valor > 0
					Group By A.Codtrab,B.CODIGO,B.MES,B.FECHAINI,B.FECHAFIN,B.FECHAPAGO)AS B            
	      WHERE TABLA1.CODTRAB=TABLA2.CODTRAB AND 
	            TABLA1.CODTRAB*=B.CODTRAB and 
	            TABLA1.NOMBOL*=B.Codigo 


         SELECT B.CODTRAB,B.CODGRUPO,A.CODNOMBOL,B.SALDO 
         FROM  Camtex.dbo.PAGOSCTA A,CAMTEX.DBO.MOVICTA B 
                        
         WHERE A.CODMOV=B.CODMOV AND 
               A.CODTRAB=B.CODTRAB
         GROUP BY B.CODTRAB,B.CODGRUPO,A.CODNOMBOL,B.SALDO




SELECT * FROM Camtex.dbo.PAGOSCTA

Select * From Camtex.dbo.nombol