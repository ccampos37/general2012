
UPDATE MAEART
SET aunidad  = 'UNI'
--SELECT * FROM MAEART
WHERE isnull(AUNIDAD,'')=''

insert MAEART
( ACODIGO,ADESCRI,AFAMILIA,AFSERIE,AFECHA,AUSER,AESTADO )

select ltrim(str(familia))+'1'+right('0000'+ltrim(str(item)),4),
DESCRIPCION,ltrim(str(familia))+'1','N','04/02/2006','carlos','V'
 from dbo.xx_inventario where not isnull(familia,'')='' AND
NOT ISNULL(item,'')=''
