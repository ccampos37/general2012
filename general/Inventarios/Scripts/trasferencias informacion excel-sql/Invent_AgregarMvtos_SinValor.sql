-- **************************************************************************
--  @NroPase= SECUENCA DE TRANSFERENCIA DE DATOS
--  Si se desea que se adicionen se debe cambiar por uno que no exista,
--  En caso contrario chancará a un paquete traslado con este Numero
-- **************************************************************************

USE nemo

declare @NroPase varchar(2)
declare @LimpiarTemporales varchar(1)
declare @FechaActualizacion varchar(10)
declare @UsuarioNombre varchar(15)


set @NroPase='Y'
set @LimpiarTemporales= 'N'

set @FechaActualizacion ='25/01/2007'
set @UsuarioNombre= 'jesus'

-- ***********************************************************
--  		Carga de Movimientos a CABECERA
-- ***********************************************************

DROP TABLE XX_CABECERA
SELECT TOP 0 * INTO XX_CABECERA FROM MOVALMCAB

---------------------------------------------------  INGRESOS
INSERT INTO XX_CABECERA(
CAALMA,CATD,CANUMDOC,CAFECDOC,
CATIPMOV,CACODMOV,CARFTDOC,CARFNDOC,
CAFECACT,CAUSUARI,CACODMON,CASITGUI,CACIERRE,CAESTIMP)
select DISTINCT
ltrim(RTRIM(Almacen)) AS ALMACEN,'NI',0,Fecha,
'I',Transaccion,LEFT(DOC_REFERENCIA,2),RTRIM(Serie)+'-'+RTRIM(Numero),
@FechaActualizacion,@UsuarioNombre,'01','V',0,@NroPase
from x_ingresos 
WHERE NOT ISNULL(ITEM,0)=0

---------------------------------------------------  SALIDAS
INSERT INTO XX_CABECERA(
CAALMA,CATD,CANUMDOC,CAFECDOC,
CATIPMOV,CACODMOV,CARFTDOC,CARFNDOC,
CAFECACT,CAUSUARI,CACODMON,CASITGUI,CACIERRE,CAESTIMP)
select DISTINCT
RTRIM(Almacen) AS ALMACEN,'NS',0,Fecha,
'S',Transaccion,LEFT(DOC_REFERENCIA,2),RTRIM(Serie)+'-'+RTRIM(Numero),
@FechaActualizacion,@UsuarioNombre,'01','V',0,@NroPase
from x_SALIDAS 
WHERE NOT ISNULL(ARTICULO,0)=0



-- ***********************************************************
--  		Carga de Movimientos a DETALLES
-- ***********************************************************
DROP TABLE XX_DETALLE
SELECT TOP 0 * INTO XX_DETALLE FROM MOVALMDET

---------------------------------------------------  INGRESOS
INSERT INTO XX_DETALLE( 
DEALMA,DETD,DENUMDOC,DEITEM,
DECODIGO,DECANTID,DEPRECIO,DECODMON,
DEDESCRI,DECANREF1,deordfab,DETR) 
select DISTINCT
RTRIM(Almacen) AS ALMACEN,'NI',0,LTRIM(STR(ITEM)) AS ITEM,
LTRIM(STR(Articulo)) AS CODIGO,tot_unidades,0,'01',
' ',0,RTRIM(Serie)+'-'+RTRIM(Numero),@NroPase
from x_ingresos 
WHERE NOT ISNULL(ITEM,0)=0


---------------------------------------------------  SALIDAS
INSERT INTO XX_DETALLE( 
DEALMA,DETD,DENUMDOC,DEITEM,
DECODIGO,DECANTID,DEPRECIO,DECODMON,
DEDESCRI,DECANREF1,deordfab,DETR) 
select DISTINCT
RTRIM(Almacen) AS ALMACEN,'NS',0,LTRIM(STR(ITEM)) AS ITEM,
LTRIM(STR(Articulo)) AS CODIGO,tot_unidades,0,'01',
' ',0,RTRIM(Serie)+'-'+RTRIM(Numero),@NroPase
from x_SALIDAS 
WHERE NOT ISNULL(ARTICULO,0)=0



-- *******************************************************************************
--  	Asigna un correlativo a cada uno de los movimientos (NI/NS) por almacen
-- *******************************************************************************
DECLARE @NUMDOC AS NVARCHAR(10)
DECLARE @CANUMDOC  AS NVARCHAR(10)
DECLARE @CAALMA AS NVARCHAR(2)
DECLARE @CATD AS NVARCHAR(2)
DECLARE @ANT_ALMA AS NVARCHAR(2)
DECLARE @ANT_TD AS NVARCHAR(2)
Declare Correla cursor for 
select CAALMA,CATD,CARFNDOC from XX_CABECERA ORDER BY CAALMA,CATD,CAFECDOC

Open Correla
fetch next from Correla into @CAALMA,@CATD,@CANUMDOC
SET @NUMDOC =case when @catd='NS' then (select tanumsal from tabalm where taalma=@caalma )
                  else (select tanument from tabalm where taalma=@caalma) end 
While @@Fetch_Status=0 
Begin 
   If @ANT_ALMA <> @CAALMA OR @ANT_TD<>@CATD 
    SET @NUMDOC =case when @catd='NS' then (select tanumsal from tabalm where taalma=@caalma )
                  else (select tanument from tabalm where taalma=@caalma) end 
   update XX_CABECERA
   set caNUMDOC=right('0000000000'+rtrim(ltrim(str(@NUMDOC +1))),10)
--   select * from
   From  XX_CABECERA A 
   Where A.CAALMA=@CAALMA AND A.CATD=@CATD AND A.CARFNDOC=@CANUMDOC  

   SET @NUMDOC = @NUMDOC+1   
   SET @ANT_ALMA = @CAALMA
   SET @ANT_TD = @CATD
   fetch next from Correla into @CAALMA,@CATD,@CANUMDOC
End
Close Correla
Deallocate Correla 


-- *******************************************************************************
-- 			Actualizando el detalle
-- *******************************************************************************
update xx_detalle
set denumdoc=a.canumdoc
--SELECT * 
    from xx_detalle b inner join xx_cabecera a
       on RTRIM(b.deordfab)=a.CARFNDOC AND B.DEALMA=A.CAALMA AND B.DETD=A.CATD



-- *******************************************************************************
-- 		Pase de informacion Temporal a las tablas definitivas
-- *******************************************************************************
DELETE MOVALMCAB WHERE CAESTIMP=@NroPase
INSERT INTO MOVALMCAB
SELECT * FROM XX_CABECERA 

--SELECT * FROM XX_CABECERA ORDER BY CAALMA, CATD,CANUMDOC 

--/* Lista elementos duplicados en "XX_CABECERA"...   Temporal de pase de informacion
--SELECT CAALMA + CATD + CANUMDOC, COUNT (CAALMA + CATD + CANUMDOC)  FROM XX_CABECERA GROUP BY CAALMA + CATD + CANUMDOC HAVING COUNT (CAALMA + CATD + CANUMDOC)>1

--/* Lista elementos de "MOVALMCAB"  que ya estan en "XX_CABECERA" (Temporal de pase de informacion)
--SELECT CAALMA + CATD + CANUMDOC 
--	FROM XX_CABECERA 
--	where (CAALMA + CATD + CANUMDOC) in (select CM.CAALMA + CM.CATD + CM.CANUMDOC from MOVALMCAB CM) 
--	ORDER BY CAALMA, CATD,CANUMDOC 

--/* Lista elementos de "XX_CABECERA" (Temporal de pase de informacion) que ya estan en MOVALMCAB
--SELECT CAALMA + CATD + CANUMDOC 
--	FROM XX_CABECERA 
--	where (CAALMA + CATD + CANUMDOC) in (select CM.CAALMA + CM.CATD + CM.CANUMDOC from MOVALMCAB CM) 
--	ORDER BY CAALMA, CATD,CANUMDOC 

DELETE MOVALMDET WHERE DETR=@NroPase
INSERT INTO MOVALMDET
SELECT * FROM XX_DETALLE 

INSERT INTO STKART (STALMA,STCODIGO) 
SELECT DISTINCT DEALMA,DECODIGO  FROM MOVALMDET
WHERE DEALMA+DECODIGO  NOT IN (SELECT DISTINCT STALMA+STCODIGO FROM STKART)
 

--SELECT * FROM X_SALIDAS
--SELECT * FROM X_INGRESOS

--DROP TABLE x_salidas
--DROP TABLE X_INGRESOS

--contador

--select caalma,catd,max(canumdoc) from movalmcab group by caalma,catd
--select * from x_salidas

--If @LimpiarTemporales= 'S'
--begin
--   DELETE X_SALIDAS
--   DELETE X_INGRESOS
--end