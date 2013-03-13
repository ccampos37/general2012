VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormAyuguia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guias"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "FormAyuguia.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   4895
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   3720
      Picture         =   "FormAyuguia.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1665
      Picture         =   "FormAyuguia.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
End
Attribute VB_Name = "FormAyuguia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Adoreg1 As ADODB.Recordset
Dim rsql As String

Private Sub Command1_Click()
If Adoreg1.RecordCount > 0 Then
    If FrmGuiaSal.TxTransa <> "" Then
            FrmGuiaSal.Text4 = Adoreg1("CANUMDOC")
    End If
    Unload Me
End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim Cod As String
AlinearAyuda Me
Init_ControlDataGrid DataGrid1

'Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.3.51;Data Source= RUTA & NAMEBD"
'DataGrid1.ClearFields                       ' Limpia las Columnas
'VGcod = "23345671"
'VGAlma = "01"
'cS1 = "Select DFCODIGO,DFDESCRI,Iif(DFSERIE = '',IiF(DFLOTE ='','  ','LOT'),'SER') as col1, "
'cS1 = cS1 & "DFCANTID,DFPRECIO,Iif(CFCODMON='US',DFIMPUS-DFIGV,DFIMPMN-DFIGV) as col2,DFIGV FROM "
'cS1 = cS1 & "FACDET A INNER JOIN FACCAB B ON A.DFTD = B.CFTD And A.DFNUMSER =B.CFNUMSER And A.DFNUMDOC=B.CFNUMDOC "
'cS1 = cS1 & "WHERE DFCODAGE = '" & cRec("CFCODAGE") & "' and DFNROCAJ = '" & cRec("CFNROCAJ") & "' and DFTD = '" & cRec("CFTD") & "' and DFNUMSER ='" & cRec("CFNUMSER") & "' and DFNUMDOC= '" & cRec("CFNUMDOC") & "'"
'Adodc1.RecordSource = cS1


rsql = "select  CATD, CANUMDOC, CAFECDOC,iif(casitgui='F','FACTURADO',IIF(CASITGUI='A','ANULADO',IIF(CASITGUI='V' or CASITGUI='P' ,'PENDIENTE',' '))) AS COL1 from MovAlmCab  where  CAALMA ='" & VGAlma & "' and CATD='GS' ORDER BY CANUMDOC"  '
Set Adoreg1 = New ADODB.Recordset
Adoreg1.Open rsql, VGcnx, adOpenDynamic, adLockOptimistic
If Adoreg1.RecordCount <> 0 Then
'?
End If
Set DataGrid1.DataSource = Adoreg1
setdata                              ' Objetos
End Sub

Private Sub setdata()        ' Carga Objetos
 'DataGrid1.Refresh
 DataGrid1.Columns(0).Locked = True
 DataGrid1.Columns(0).WrapText = True
 'DataGrid1.Columns(0). = "CATD"
 DataGrid1.Columns(0).Caption = "   TD"
 DataGrid1.Columns(1).Caption = "  NUMERO"
 DataGrid1.Columns(2).Caption = "   FECHA"
  DataGrid1.Columns(3).Caption = " SITUACION"
  DataGrid1.Columns(0).Width = 800
  DataGrid1.Columns(1).Width = 1500
  DataGrid1.Columns(2).Width = 1500
  DataGrid1.Columns(3).Width = 1500
 'DataGrid1.Columns(0).WrapText = False
End Sub
