create   function fn_unnumerorang(@num bigint,@range varchar(200)) 
returns bigint
as
Begin
declare  @i bigint,@cadaux varchar(200),@cont as bigint,@valor bigint,
         @r1 as bigint,@r2 as bigint
set @i=1
set @cadaux=''
set @cont=0
set @valor=0
While @i <=len(@range)
begin	
	if substring(@range,@i,1)<>',' set @cadaux=@cadaux+substring(@range,@i,1)		
	if substring(@range,@i,1)=',' 
    begin		
		--print @cadaux
		set @cont=@cont+1
		if @cont=1 set @r1=cast(@cadaux as bigint)
        if @cont=2 
		begin 
			set @r2=cast(@cadaux as bigint)
			set @valor=case when @num between @r1 and @r2 then @r1 else @valor end 			            
            set @cont=1              
            set @r1=@r2 
            set @r2=0
		end 
		--print @r1
		set @cadaux=''        
    end  	
	set @i=@i+1
end
return @valor
end


fn_unnumerorang



Declare 
@num bigint,@range varchar(200)

set @num=80
set @range='70,80,90,100,1000,'


declare  @i bigint,@cadaux varchar(200),@cont as bigint,@valor bigint,
         @r1 as bigint,@r2 as bigint

set @i=1
set @cadaux=''
set @cont=0
set @valor=0
While @i <=len(@range)
begin	
	if substring(@range,@i,1)<>',' set @cadaux=@cadaux+substring(@range,@i,1)		
	if substring(@range,@i,1)=',' 
    begin		
		--print @cadaux
		set @cont=@cont+1
		if @cont=1 set @r1=cast(@cadaux as bigint)
        if @cont=2 
		begin 
			set @r2=cast(@cadaux as bigint)
			set @valor=case when @num between @r1 and @r2 then @r1 else @valor end 			            
            set @cont=1              
            set @r1=@r2 
            set @r2=0
		end 
		--print @r1
		set @cadaux=''        
    end  	
	set @i=@i+1
end
Print @valor