VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmAsientoPrevio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asiento Contable"
   ClientHeight    =   6420
   ClientLeft      =   930
   ClientTop       =   600
   ClientWidth     =   9135
   Icon            =   "FrmAsientoPrevio.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9135
   Begin VB.CommandButton Command2 
      Caption         =   "&Enviar"
      Height          =   360
      Left            =   72
      TabIndex        =   7
      Top             =   6012
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   1512
      TabIndex        =   6
      Top             =   6012
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Costo Ventas"
      Height          =   2892
      Left            =   72
      TabIndex        =   3
      Top             =   3096
      Width           =   9012
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   1512
         Left            =   108
         TabIndex        =   4
         Top             =   1296
         Width           =   8784
         _ExtentX        =   15505
         _ExtentY        =   2672
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         Caption         =   "DETALLE"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Dmov_Secue"
            Caption         =   "Secuencial"
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
            DataField       =   "dmov_cuent"
            Caption         =   "       Cuenta"
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
         BeginProperty Column02 
            DataField       =   "dmov_fecha"
            Caption         =   "       Fecha"
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
         BeginProperty Column03 
            DataField       =   "dMov_Glosa"
            Caption         =   "         Glosa"
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
         BeginProperty Column04 
            DataField       =   "dmov_debe"
            Caption         =   "          Debe"
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
         BeginProperty Column05 
            DataField       =   "dmov_Haber"
            Caption         =   "        Haber"
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
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2280.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1365.165
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   984
         Left            =   108
         TabIndex        =   5
         Top             =   288
         Width           =   8772
         _ExtentX        =   15478
         _ExtentY        =   1746
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         Caption         =   "CABEZERA"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "subdiar_codigo"
            Caption         =   "Sub Diario"
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
            DataField       =   "cmov_c_compr"
            Caption         =   "Comprobante"
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
         BeginProperty Column02 
            DataField       =   "cmov_Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column03 
            DataField       =   "cmov_glosa"
            Caption         =   "Glosa"
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
         BeginProperty Column04 
            DataField       =   "cmov_moned"
            Caption         =   "Moneda"
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
         BeginProperty Column05 
            DataField       =   "Cmov_tipca"
            Caption         =   "Tipo Cambio"
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
               Alignment       =   2
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1035.213
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Salidas"
      Height          =   2928
      Left            =   72
      TabIndex        =   0
      Top             =   108
      Width           =   9012
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1476
         Left            =   108
         TabIndex        =   1
         Top             =   1368
         Width           =   8784
         _ExtentX        =   15505
         _ExtentY        =   2619
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         Caption         =   "DETALLE"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Dmov_Secue"
            Caption         =   "Secuencial"
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
            DataField       =   "dmov_cuent"
            Caption         =   "       Cuenta"
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
         BeginProperty Column02 
            DataField       =   "dmov_fecha"
            Caption         =   "       Fecha"
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
         BeginProperty Column03 
            DataField       =   "dMov_Glosa"
            Caption         =   "         Glosa"
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
         BeginProperty Column04 
            DataField       =   "dmov_debe"
            Caption         =   "          Debe"
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
         BeginProperty Column05 
            DataField       =   "dmov_Haber"
            Caption         =   "        Haber"
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
               Locked          =   -1  'True
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2280.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1365.165
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   984
         Left            =   108
         TabIndex        =   2
         Top             =   360
         Width           =   8772
         _ExtentX        =   15478
         _ExtentY        =   1746
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         Caption         =   "CABEZERA"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "subdiar_codigo"
            Caption         =   "Sub Diario"
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
            DataField       =   "cmov_c_compr"
            Caption         =   "Comprobante"
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
         BeginProperty Column02 
            DataField       =   "cmov_Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column03 
            DataField       =   "cmov_glosa"
            Caption         =   "Glosa"
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
         BeginProperty Column04 
            DataField       =   "cmov_moned"
            Caption         =   "Moneda"
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
         BeginProperty Column05 
            DataField       =   "Cmov_tipca"
            Caption         =   "Tipo Cambio"
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
               Alignment       =   2
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3435.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1035.213
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmAsientoPrevio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cConexAux As ADODB.Connection
Dim cVGDBT  As ADODB.Connection
Dim csql As String
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim Adodc3 As ADODB.Recordset
Dim Adodc4 As ADODB.Recordset

Dim cSel1 As New ADODB.Recordset
Dim cSel2 As New ADODB.Recordset

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo Err

csql = "insert into CabMov" & Format(FrmAsiento2001.Mes1, "00") & " in '" + VGParamSistem.RutaReport & VGContTra & Year(Date) & ".MDB" + "' select * from CabMov1"
cConexAux.Execute csql
csql = "insert into DetMov" & Format(FrmAsiento2001.Mes1, "00") & " in '" + VGParamSistem.RutaReport & VGContTra & Year(Date) & ".MDB" + "' select * from DetMov1"
cConexAux.Execute csql

csql = "Update al_Kardex_Val set Asiento=true"
cConexAux.Execute csql


csql = "Select * From al_Kardex_Val"
cSel1.Open csql, cConexAux, adOpenStatic
Do While Not cSel1.EOF
    If Not IsNull(cSel1("Tip_Transa")) Then
       csql = "Select asiento From MovAlmCab where catd='" & cSel1("tip_transa") & "' and canumdoc='" & cSel1("Num_Doc") & "' and caalma='" & VGAlma & "'"
       cSel2.Open csql, VGCNx, adOpenStatic
       If Not cSel2.EOF Then
              csql = "Update MovAlmCab set Asiento=true where catd='" & cSel1("tip_transa") & "' and canumdoc='" & cSel1("Num_Doc") & "' and caalma='" & VGAlma & "'"
              VGCNx.Execute csql
       End If
       cSel2.Close
    End If
    cSel1.MoveNext
Loop
cSel1.Close

cConexAux.Execute "Delete From CabMov1"
cConexAux.Execute "Delete From DetMov1"

'Adodc2.Close
SQL = "Select * From CabMov1"
Set Adodc2 = New ADODB.Recordset
Adodc2.Open SQL, cConexAux, adOpenStatic
SQL = "Select * From DetMov1"
'Adodc1.Close
Set adodc1 = New ADODB.Recordset
adodc1.Open SQL, cConexAux, adOpenStatic
Set DataGrid2.DataSource = Adodc2
Set DataGrid1.DataSource = adodc1

Set DataGrid4.DataSource = Adodc2
Set DataGrid3.DataSource = adodc1

DataGrid2.Refresh
DataGrid1.Refresh

DataGrid4.Refresh
DataGrid3.Refresh

Command2.Enabled = False
FrmAsiento.Command1.Enabled = False
Exit Sub
Err:
  If Err.Number = -2147467259 Then
     
     MsgBox "El numero de comprobante ya fue utilizado en contabilidad," & Chr(13) & "verifique en contabilidad el nro comprobante ", vbInformation, "Aviso"
  Else
     MsgBox Err.Description, vbInformation
  End If
End Sub

Private Sub Command3_Click()
On Error GoTo Err
'RMM********************ingresos*****************
csql = "insert into CabMov" & Format(FrmAsiento.Mes1, "00") & " in '" + VGParamSistem.RutaReport & VGContTra & Year(Date) & ".MDB" + "' select * from CabMov1"
cConexAux.Execute csql
csql = "insert into DetMov" & Format(FrmAsiento.Mes1, "00") & " in '" + VGParamSistem.RutaReport & VGContTra & Year(Date) & ".MDB" + "' select * from DetMov1"
cConexAux.Execute csql

csql = "Update al_Kardex_Val set Asiento=true"
cConexAux.Execute csql


csql = "Select * From al_Kardex_Val"
cSel1.Open csql, cConexAux, adOpenStatic
Do While Not cSel1.EOF
    If Not IsNull(cSel1("Tip_Transa")) Then
       csql = "Select asiento From MovAlmCab where catd='" & cSel1("tip_transa") & "' and canumdoc='" & cSel1("Num_Doc") & "' and caalma='" & VGAlma & "'"
       cSel2.Open csql, VGCNx, adOpenStatic
       If Not cSel2.EOF Then
              csql = "Update MovAlmCab set Asiento=true where catd='" & cSel1("tip_transa") & "' and canumdoc='" & cSel1("Num_Doc") & "' and caalma='" & VGAlma & "'"
              VGCNx.Execute csql
       End If
       cSel2.Close
    End If
    cSel1.MoveNext
Loop
cSel1.Close

cConexAux.Execute "Delete From CabMov1"
cConexAux.Execute "Delete From DetMov1"

Adodc2.Close
SQL = "Select * From CabMov1"
Adodc2.Open SQL, cConexAux, adOpenStatic
SQL = "Select * From DetMov1"
adodc1.Close
adodc1.Open SQL, cConexAux, adOpenStatic
Set DataGrid2.DataSource = Adodc2
Set DataGrid1.DataSource = adodc1
DataGrid2.Refresh
DataGrid1.Refresh
Command2.Enabled = False
FrmAsiento.Command1.Enabled = False
Exit Sub
Err:
  If Err.Number = -2147467259 Then
     MsgBox "El numero de comprobante ya fue utilizado en contabilidad," & Chr(13) & "verifique en contabilidad el nro comprobante ", vbInformation, "Aviso"
  Else
     MsgBox Err.Description, vbInformation
  End If

End Sub

Private Sub DataGrid2_Click()
If Not Adodc2.EOF Then
   adodc1.Close
   SQL = "Select * From DetMov1 where SubDiar_Codigo='" & Adodc2("SubDiar_Codigo") & "' and Dmov_C_Compr='" & Adodc2("CMov_C_Compr") & "'"
   adodc1.Open SQL, cConexAux, adOpenStatic
   If adodc1.BOF Or adodc1.EOF Then
      Exit Sub
   End If
   adodc1.MoveFirst
   Set DataGrid1.DataSource = adodc1
   DataGrid1.Refresh
End If
End Sub

Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If Not Adodc2.EOF Then
   adodc1.Close
   SQL = "Select * From DetMov1 where SubDiar_Codigo='" & Adodc2("SubDiar_Codigo") & "' and Dmov_C_Compr='" & Adodc2("CMov_C_Compr") & "'"
   adodc1.Open SQL, cConexAux, adOpenStatic
   If adodc1.BOF Or adodc1.EOF Then
      Exit Sub
   End If
   adodc1.MoveFirst
   Set DataGrid1.DataSource = adodc1
   DataGrid1.Refresh
End If
End Sub

Private Sub Form_Load()
Set Adodc2 = New ADODB.Recordset
Set adodc1 = New ADODB.Recordset
Set Adodc3 = New ADODB.Recordset
Set Adodc4 = New ADODB.Recordset

central Me
Conectar
SQL = "Select * From CabMov1 where SubDiar_Codigo='" & FrmAsiento2001.subdiar & "'"
Adodc2.Open SQL, cConexAux, adOpenStatic
If Not Adodc2.BOF And Not Adodc2.EOF Then
   
   Adodc2.MoveFirst
   If Not Adodc2.EOF Then
      SQL = "Select * From DetMov1 where SubDiar_Codigo='" & Adodc2("SubDiar_Codigo") & "' and Dmov_C_Compr='" & Adodc2("CMov_C_Compr") & "'"
      adodc1.Open SQL, cConexAux, adOpenStatic
      If adodc1.BOF Or adodc1.EOF Then
         Exit Sub
      End If
      adodc1.MoveFirst
      Set DataGrid1.DataSource = adodc1
      DataGrid1.Refresh
   End If
   Set DataGrid2.DataSource = Adodc2
   DataGrid2.Refresh
End If

SQL = "Select * From CabMov1 where SubDiar_Codigo='" & FrmAsiento2001.SubdiarCompra & "'"
Adodc4.Open SQL, cConexAux, adOpenStatic
If Not Adodc4.BOF And Not Adodc4.EOF Then
    Adodc4.MoveFirst
    If Not Adodc4.EOF Then
       SQL = "Select * From DetMov1 where SubDiar_Codigo='" & Adodc4("SubDiar_Codigo") & "' and Dmov_C_Compr='" & Adodc4("CMov_C_Compr") & "'"
       Adodc3.Open SQL, cConexAux, adOpenStatic
       If Adodc3.BOF Or Adodc3.EOF Then
          Exit Sub
       End If
       Adodc3.MoveFirst
       Set DataGrid3.DataSource = Adodc3
       DataGrid3.Refresh
    End If
    Set DataGrid4.DataSource = Adodc4
    DataGrid4.Refresh
    
End If

End Sub

Sub Conectar()
cRt = App.Path & "\BdAuxCom.Mdb"
Set cConexAux = New ADODB.Connection
cConexAux.CursorLocation = adUseClient
cConexAux.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRt & ";"
cConexAux.Open

If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
   Set cVGDBT = New ADODB.Connection
   
   cVGDBT.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & VGParamSistem.RutaReport & VGContTra & Year(Date) & ".MDB;"
   ' With cVGDBT 'para Movimientos
   '     .CursorLocation = adUseClient
   '     .Provider = "Microsoft.Jet.OLEDB.3.51"
   '     .ConnectionString = "Data Source=" & VGParamSistem.RutaReport & VGContTra & Year(Date) & ".MDB"
    cVGDBT.Open
   'End With
End If
End Sub

