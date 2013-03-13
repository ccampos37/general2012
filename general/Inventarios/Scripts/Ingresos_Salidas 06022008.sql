USE invnemoplanta
declare @NroPase varchar(2)
set @NroPase='A'


--SELECT * FROM XX_CABECERA
DROP TABLE XX_CABECERA
SELECT * INTO XX_CABECERA FROM MOVALMCAB
DELETE  XX_CABECERA

--SELECT * FROM XX_CABECERA

INSERT INTO XX_CABECERA
(
CAALMA,CATD,CANUMDOC,CAFECDOC,
CATIPMOV,CACODMOV,CARFTDOC,CARFNDOC,
CAFECACT,CAUSUARI,CACODMON,CASITGUI,CACIERRE,CAESTIMP
)

select DISTINCT
ltrim(RTRIM(Almacen)) AS ALMACEN,'NI',0,Fecha,
'I',Transaccion,LEFT(DOC_REFERENCIA,2),RTRIM(Serie)+'-'+RTRIM(Numero),
'01/01/2008','JESUS','01','V',
0,@NroPase
--Destino,T/C,
--REPRESENTANTE,
--TOTAL_MANOJOS,
--factor,

from x_ingresos 
WHERE NOT ISNULL(ITEM,0)=0

INSERT INTO XX_CABECERA
(
CAALMA,CATD,CANUMDOC,CAFECDOC,
CATIPMOV,CACODMOV,CARFTDOC,CARFNDOC,
CAFECACT,CAUSUARI,CACODMON,CASITGUI,CACIERRE,CAESTIMP
)

select DISTINCT
RTRIM(Almacen) AS ALMACEN,'NS',0,Fecha,
'S',Transaccion,LEFT(DOC_REFERENCIA,2),RTRIM(Serie)+'-'+RTRIM(Numero),
'01/01/2008','JESUS','01','V',
0,@NroPase
--Destino,T/C,
--REPRESENTANTE,
--TOTAL_MANOJOS,
--factor,

from x_SALIDAS 
WHERE NOT ISNULL(ARTICULO,0)=0



--SELECT * FROM XX_CABECERA ORDER BY 1

--- DETALLE DE DOCUMENTOS



DROP TABLE XX_DETALLE
SELECT * INTO XX_DETALLE FROM MOVALMDET
dELETE  xX_DETALLE

--INGRESOS

INSERT INTO XX_DETALLE 
( 
DEALMA,DETD,DENUMDOC,DEITEM,
DECODIGO,DECANTID,DEPRECIO,DECODMON,
DEDESCRI,DECANREF1,deordfab,DETR,decodmov) 

select DISTINCT
RTRIM(Almacen) AS ALMACEN,'NI',0,LTRIM(STR(ITEM)) AS ITEM,
--Fecha,Transaccion,doc_referencia,Serie,Numero,Destino,
--T/C,
LTRIM(STR(Articulo)) AS CODIGO,tot_unidades,mvtprcunit,'01',
' ',0,RTRIM(Serie)+'-'+RTRIM(Numero),@NroPase,Transaccion
--'B'
--REPRESENTANTE,
--TOTAL_MANOJOS,
--factor,

from x_ingresos 
WHERE NOT ISNULL(ITEM,0)=0


--SALIDAS

INSERT INTO XX_DETALLE 
( 
DEALMA,DETD,DENUMDOC,DEITEM,
DECODIGO,DECANTID,DEPRECIO,DECODMON,
DEDESCRI,DECANREF1,deordfab,DETR,decodmov) 

--SELECT * FROM XXX1

select DISTINCT
RTRIM(Almacen) AS ALMACEN,'NS',0,LTRIM(STR(ITEM)) AS ITEM,
--Fecha,Transaccion,doc_referencia,Serie,Numero,Destino,
--T/C,
LTRIM(STR(Articulo)) AS CODIGO,tot_unidades,0,'01',
' ',0,RTRIM(Serie)+'-'+RTRIM(Numero),@NroPase, Transaccion
--REPRESENTANTE,
--TOTAL_MANOJOS,
--factor,

from x_SALIDAS 
WHERE NOT ISNULL(ARTICULO,0)=0


DECLARE @NUMDOC AS NVARCHAR(10)
DECLARE @CANUMDOC  AS NVARCHAR(10)
DECLARE @CANUMDOC1  AS NVARCHAR(10)
DECLARE @CAALMA AS NVARCHAR(2)
DECLARE @CATD AS NVARCHAR(2)
declare @Codmov as nvarchar(2)
DECLARE @ANT_ALMA AS NVARCHAR(2)
DECLARE @ANT_TD AS NVARCHAR(2)
declare @ant_codmov as nvarchar(2)
DECLARE @ANT_CARFNDOC AS NVARCHAR(10)

Declare Correla cursor for 
select CAALMA,CATD,CARFNDOC,cacodmov from XX_CABECERA ORDER BY CAALMA,CATD,CAFECDOC

Open Correla
fetch next from Correla into @CAALMA,@CATD,@CANUMDOC1,@codmov
SET @NUMDOC =case when @catd='NS' then (select tanumsal from tabalm where taalma=@caalma )
                  else (select tanument from tabalm where taalma=@caalma) end 
--select tanument from tabalm where taalma='01'
--select tanumsal from tabalm where taalma='01'
SET @ANT_ALMA = @CAALMA
SET @ANT_TD = @CATD
set @ANT_CARFNDOC = @CANUMDOC1
set @ant_codmov=@codmov
While @@Fetch_Status=0 
Begin 
   If @ANT_ALMA <> @CAALMA OR @ANT_TD<>@CATD OR @ANT_CARFNDOC <> @CANUMDOC1 or @ant_codmov <> @codmov
	begin
      SET @NUMDOC =case when @catd='NS' then (select tanumsal from tabalm where taalma=@caalma )
                        else (select tanument from tabalm where taalma=@caalma) end
	  print @NUMDOC + ' - ' + @catd
	end
   update XX_CABECERA
   set caNUMDOC=right('0000000000'+rtrim(ltrim(str(@NUMDOC +1))),11)
	-- select * from
   From  XX_CABECERA A 
   Where A.CAALMA=@CAALMA AND A.CATD=@CATD and a.cacodmov=@codmov AND A.CARFNDOC=@CANUMDOC1  
	-- select * from xx_cabecera
   SET @NUMDOC = @NUMDOC+1   
   SET @ANT_ALMA = @CAALMA
   SET @ANT_TD = @CATD
   set @ANT_CARFNDOC = @CANUMDOC1
   set @ant_codmov=@codmov

   -- Actualizo contadores tabla TABALM

   if @CATD = 'NS'
      begin
	update TABALM
	set tanumsal = @NUMDOC
	where taalma = @caalma
	print 'actualizando NS'
      end
   else
      begin
	update TABALM
	set tanument = @NUMDOC
	where taalma = @caalma
	print 'actualizando NI'
      end

   fetch next from Correla into @CAALMA,@CATD,@CANUMDOC1,@codmov


End
Close Correla
Deallocate Correla 

--- actualizando el detalle


update xx_detalle
set denumdoc=a.canumdoc
--SELECT * 
    from xx_detalle b inner join xx_cabecera a
       on RTRIM(b.deordfab)=a.CARFNDOC AND B.DEALMA=A.CAALMA AND B.DETD=A.CATD and b.decodmov=a.cacodmov


--SELECT * FROM XX_CABECERA ORDER BY 1,2,4,9
--select dealma,detd,max(denumdoc) from movalmdet  group by dealma,detd 
--select * from movalmdet  order by dealma,detd,denumdoc desc 
---SELECT * FROM XX_DETALLE INNER JOIN MAEART ON RTRIM(DECODIGO)=ACODIGO WHERE DEALMA='08'
--select * from x_salidas


DELETE MOVALMCAB WHERE CAESTIMP=@NroPase
INSERT INTO MOVALMCAB
SELECT * FROM XX_CABECERA 

--SELECT * FROM XX_CABECERA
--select * from movalmcab  ORDER BY CAALMA, CATD,CANUMDOC 

--/* Lista elementos duplicados en "XX_CABECERA"...   Temporal de pase de informacion
--SELECT CAALMA + CATD + CANUMDOC , COUNT (CAALMA + CATD + CANUMDOC)  FROM XX_CABECERA GROUP BY CAALMA + CATD + CANUMDOC HAVING COUNT (CAALMA + CATD + CANUMDOC )>1
--SELECT CARFNDOC,* From  XX_CABECERA Where CANUMDOC='0          ' ORDER BY CANUMDOC




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


--select serie,numero,articulo,fecha from x_ingresos order by serie,numero,articulo,fecha
--select serie,numero,articulo from x_ingresos
--Encontrar documentos duplicados
--select serie+ltrim(str(numero))+articulo from x_ingresos group by serie+ltrim(str(numero))+articulo having count(serie+ltrim(str(numero))+articulo)>1


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
--DROP TABLE Xx_cabecera 
--DROP TABLE xx_detalle

--contador

--select caalma,catd,max(canumdoc) from movalmcab group by caalma,catd
--select * from x_salidas

--DELETE X_SALIDAS

--SELECT * FROM MOVALMCAB WHERE CAALMA='06' ORDER BY CAESTIMP
--DELETE X_INGRESOS


--SELECT * FROM X_INGRESOS