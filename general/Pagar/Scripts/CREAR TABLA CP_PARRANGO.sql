drop table CP_PARRANGO
CREATE TABLE cp_rangovcto(COD BIGINT,DESCRIP VARCHAR(100)) 


INSERT INTO cp_rangovcto
SELECT 0,'- 7 Dias' union all 
SELECT 7,'De 7 a 14 Dias   ' union all 
SELECT 14,'De 14 a 30 Dias ' union all 
SELECT 30,'De 30 a 60 Dias ' union all 
SELECT 60,'De 60 a 90 Dias ' union all 
SELECT 90,'+ 90 Dias ' 