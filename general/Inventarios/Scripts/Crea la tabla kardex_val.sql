/*Añadir Tabla: Kardex_val a la Base Datos*/

if (select count(name) from sysobjects where name like 'kardex_val')>0 
	drop table kardex_val

go	

CREATE TABLE [Kardex_Val] (
	[COD_ART] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[FEC_DOC] [smalldatetime] NULL ,
	[HOR_DOC] [smalldatetime] NULL ,
	[COD_MOV] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL,
	[TIP_TRANSA] [varchar](2) COLLATE Modern_Spanish_CI_AS NULL,
	[NUM_DOC] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL,
	[CAN_ART] [float],
	[PRE_UNIT] [float],
	[COS_PRO] [float],
	[SAL_STOCK] [float],
	[SER_LOT] [float],
	[ING_SAL] [float]		
) ON [PRIMARY]



