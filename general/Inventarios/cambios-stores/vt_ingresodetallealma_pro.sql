alter    PROC vt_ingresodetallealma_pro
@base varchar(50),
@tabla varchar(50),
@tipo char(1),
@item char(3),
@numero char(11),
@producto varchar(8),
@unidad char(3),
@cantidad  float,
@preciopacto float,
@dsctoxitem float,
@importebruto float,
@porcomision float,
@mdsctoitem float,
@mdsctoxlinea float,
@mdsctoxprom float,
@mimpor float,
@unidadref float,
@almacen varchar(2)
AS
Declare @cadena nvarchar(1000)
Declare @parame nvarchar(1000)
Declare @valor varchar(2)
Declare @GS varchar(2)

set @valor=@almacen
if @tipo='1' 
  begin 
   set @gs='GS'
  end 
if @tipo='2'   --nota salida
  begin 
   set @gs='NS'
  end 
if @tipo='3'  --nota ingreso
  begin 
   set @gs='NI'
  end 



      Set @cadena='INSERT INTO ['+@base +'].dbo.'+@tabla+
		  '  (DETD,
			DEALMA,
			DEITEM,
			DENUMDOC,
			DECODIGO,
			DECANTID,
			DEPREVTA,
			DEDESCTO,
			DEVALTOT,
			DEPORDES,
			DECANTENT,
			DECANREF1
			)
		VALUES(
			@GS,
			@valor,
			@item,
			@numero,
			@producto,
			@cantidad,
			@preciopacto,
			@dsctoxitem,
			@importebruto,
			@mdsctoitem,			
			@cantidad,
			@unidadref)'

	Set @parame=N'@item char(3),
		@numero char(11),
		@producto char(8),
		@unidad char(3),
		@cantidad  float,
		@preciopacto float,
		@dsctoxitem float,
		@importebruto float,
		@porcomision float,
		@mdsctoitem float,
		@mdsctoxlinea float,
		@mdsctoxprom float,
		@mimpor float,
		@unidadref float,
		@valor varchar(2),
		@GS varchar(2)'

       execute sp_executesql @cadena,@parame,@item,
						@numero,
						@producto,
						@unidad,
						@cantidad,
						@preciopacto,
						@dsctoxitem,
						@importebruto,
						@porcomision,
						@mdsctoitem,
						@mdsctoxlinea,
						@mdsctoxprom,
						@mimpor,
						@unidadref,
						@valor,
						@gs
	


























