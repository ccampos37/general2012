set nocount on
Declare @documentocargo varchar(2)
declare @cargoapeimpape float
Declare @cargonumdoc varchar(15)
Declare @clientecodigo varchar(20)

DECLARE tablas CURSOR FOR 
--update cp_cargo set cargoapeimppag=0,cargoapeflgcan='0'
SELECT documentocargo,cargonumdoc,cargoapeimpape,clientecodigo from cp_cargo WHERE cargoapeflgcan='0'
--	cargonumdoc='00100050298' 

	OPEN tablas
	/* Leer cada registro del cursor  */
	FETCH NEXT FROM tablas INTO @documentocargo,@cargonumdoc,@cargoapeimpape,@clientecodigo

	WHILE @@FETCH_STATUS = 0
	BEGIN
        	declare @cadsql varchar(3000)
		
		--print @cargonumdoc
		--print @cargoapeimpape	

		set @cadsql=''
		set @cadsql='update cp_cargo 
	           set cargoapeimppag=(select isnull(sum(isnull(abonocanimpsol,0)),0) from cp_abono 
		   	where documentoabono=''' +@documentocargo+ ''' and abonocancli=''' +@clientecodigo+ ''' and abononumdoc=''' +@cargonumdoc+''')
		   where documentocargo=''' +@documentocargo+ ''' and clientecodigo=''' +@clientecodigo+ ''' and cargonumdoc=''' +@cargonumdoc+ ''''
		exec(@cadsql)
				
		set @cadsql=''
		set @cadsql='if round( cast( (select isnull(sum(isnull(abonocanimpsol,0)),0) from cp_abono 
		  where documentoabono=''' +@documentocargo+ ''' and abonocancli=''' +@clientecodigo+ ''' and abononumdoc=''' +@cargonumdoc +''') as varchar(15)),2) >=' + cast(@cargoapeimpape as varchar(15)) + 
			'  begin
			   update cp_cargo set cargoapeflgcan=''1'' where documentocargo=''' +@documentocargo+ ''' and clientecodigo=''' +@clientecodigo+ ''' and cargonumdoc=''' +@cargonumdoc+ ''' end'
		exec(@cadsql)
		

 	FETCH NEXT FROM tablas INTO @documentocargo,@cargonumdoc,@cargoapeimpape,@clientecodigo
    END
	CLOSE tablas
	DEALLOCATE tablas

set nocount off
--select * from cp_abono
--select * from cp_cargo
