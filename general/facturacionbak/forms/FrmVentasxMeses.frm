VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form FrmVentasxMeses 
   Caption         =   "Ventas Anuales x Zona"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   ScaleHeight     =   6750
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5955
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   10504
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmVentasxMeses.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNumReg"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataGrid1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmbotones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmVentasxMeses.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cCancela"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cAcepta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame2 
         Caption         =   "Meses"
         Height          =   3495
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   6495
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   8
            Left            =   1560
            TabIndex        =   15
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   16
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   5
            Left            =   1560
            TabIndex        =   17
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   18
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   7
            Left            =   1560
            TabIndex        =   19
            Top             =   2280
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   20
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   1
            Left            =   4800
            TabIndex        =   21
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   2
            Left            =   4800
            TabIndex        =   22
            Top             =   2760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   9
            Left            =   4800
            TabIndex        =   23
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   10
            Left            =   4800
            TabIndex        =   24
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   11
            Left            =   4800
            TabIndex        =   25
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            SignoNegativo   =   0   'False
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin TextFer.TxFer txt 
            Height          =   375
            Index           =   12
            Left            =   4800
            TabIndex        =   26
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            TipoDato        =   1
            MarcarTextoAlEnfoque=   -1  'True
         End
         Begin VB.Label lbl 
            Caption         =   "Febrero"
            Height          =   285
            Index           =   5
            Left            =   240
            TabIndex        =   38
            Top             =   930
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Marzo"
            Height          =   285
            Index           =   6
            Left            =   240
            TabIndex        =   37
            Top             =   1395
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Abril"
            Height          =   285
            Index           =   7
            Left            =   240
            TabIndex        =   36
            Top             =   1980
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Mayo"
            Height          =   285
            Index           =   8
            Left            =   240
            TabIndex        =   35
            Top             =   2415
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Junio"
            Height          =   285
            Index           =   9
            Left            =   240
            TabIndex        =   34
            Top             =   2895
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Enero"
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   33
            Top             =   450
            Width           =   705
         End
         Begin VB.Label lbl 
            Caption         =   "Agosto"
            Height          =   285
            Index           =   0
            Left            =   3360
            TabIndex        =   32
            Top             =   930
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Setiembre"
            Height          =   285
            Index           =   1
            Left            =   3360
            TabIndex        =   31
            Top             =   1515
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Octubre"
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   30
            Top             =   1980
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Noviembre"
            Height          =   285
            Index           =   3
            Left            =   3360
            TabIndex        =   29
            Top             =   2415
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Diciembre"
            Height          =   285
            Index           =   10
            Left            =   3360
            TabIndex        =   28
            Top             =   2895
            Width           =   780
         End
         Begin VB.Label lbl 
            Caption         =   "Julio"
            Height          =   285
            Index           =   11
            Left            =   3360
            TabIndex        =   27
            Top             =   450
            Width           =   705
         End
      End
      Begin VB.Frame frmbotones 
         Height          =   555
         Left            =   -74520
         TabIndex        =   5
         Top             =   4830
         Width           =   5730
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Nuevo"
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   10
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "E&ditar"
            Height          =   330
            Index           =   1
            Left            =   1185
            TabIndex        =   9
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   2
            Left            =   2310
            TabIndex        =   8
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Salir"
            Height          =   330
            Index           =   4
            Left            =   4560
            TabIndex        =   7
            Top             =   165
            Width           =   1080
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Imprimir"
            Height          =   330
            Index           =   3
            Left            =   3435
            TabIndex        =   6
            Top             =   165
            Width           =   1080
         End
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   90
         TabIndex        =   3
         Top             =   420
         Width           =   6570
         Begin TextFer.TxFer txt 
            Height          =   420
            Index           =   0
            Left            =   1455
            TabIndex        =   4
            Top             =   195
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   741
            Object.CausesValidation=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   3
            Text            =   ""
            ColorIlumina    =   -2147483624
            SaltarAlEnter   =   -1  'True
            Valor           =   ""
            NoCaracteres    =   "0123456789"
            MarcarTextoAlEnfoque=   -1  'True
            NoRangoCadena   =   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Periodo :"
            Height          =   270
            Left            =   600
            TabIndex        =   13
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   585
         Left            =   2025
         TabIndex        =   2
         Top             =   5055
         Width           =   1380
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   585
         Left            =   3945
         TabIndex        =   1
         Top             =   5055
         Width           =   1380
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3255
         Left            =   -74760
         TabIndex        =   39
         Top             =   840
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5741
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
               LCID            =   10250
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
               LCID            =   10250
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
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -69315
         TabIndex        =   12
         Top             =   4305
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Registros"
         Height          =   270
         Left            =   -70260
         TabIndex        =   11
         Top             =   4320
         Width           =   900
      End
   End
End
Attribute VB_Name = "FrmVentasxMeses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim rs As New ADODB.Recordset
Dim rsAsiento As ADODB.Recordset

Private Sub ChkCargo_Click()
    If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatosAsiento
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rsAsiento = Nothing
  Set VGvardllgen = Nothing
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
  Me.Width = 7050
  Me.Height = 6255
End Sub

'FIXIT: Declare 'MuestraDatosAsiento' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Public Function MuestraDatosAsiento()
 Dim SQL As String
  
  SQL = "SELECT A.zonacodigo, A.zonadescripcion,,"
  SQL = SQL & "B.zomames01, B.zomames02, B.zomames03, B.zomames04,"
  SQL = SQL & "B.zomames05, B.zomames06, B.zomames07, B.zomames08,"
  SQL = SQL & "B.zomames09, B.zomames10, B.zomames11, B.zomames12,"
  SQL = SQL & "FROM  vt_zona a,vt_ventasxzonas A "
  SQL = SQL & "and b.zonacodigo='" & VGParametros.empresacodigo & "'"
  SQL = SQL & " and A.zonacodigo=B.zonacodigo and B.anno='" & VGParamSistem.AnoProceso & "'"
  
  Set rs = VGCNx.Execute(SQL)
  Set DataGrid1.DataSource = rs
  Call ConfiguraGridAsientos
  lblNumReg.Caption = rs.RecordCount
  
End Function

Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim spos As Integer
  Dim SQL As String
  Dim SQL1 As String
  
  On Error GoTo X
  
  Select Case Index
     Case 0   'nuevo
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
        txt(0).SetFocus
        Call ModoEditable(True, FrmVentasxMeses, "")
        frmbotones.Visible = False
        modoinsert = True
        
     Case 1   'modificar
        If DataGrid1.Row < 0 Then
          Exit Sub
        End If
        SSTab1.Tab = 1
        SSTab1.TabEnabled(1) = True
        modoedit = True
        frmbotones.Visible = False
        Call EditarAsiento
        Call ModoEditable(True, FrmVentasxMeses, "")
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro de SubAsiento Nº " & TDBGridAsientos.Columns(0).Value & "?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM CT_ASIENTO WHERE asientocodigo='" & txt(0).Text & "'"
          SQL1 = "DELETE FROM CT_ASIENTOCorre WHERE "
          SQL1 = SQL1 & " empresacodigo='" & VGParametros.empresacodigo & "' and asientocodigo='" & txt(0).Text & "' AND asientoanno='" & VGParamSistem.AnoProceso & "'"
          VGCNx.Execute (SQL)
          VGCNx.Execute (SQL1)
          Call MuestraDatosAsiento
       End If
        
     Case 3   'imprimir
       Call Impresion("rptAsiento.rpt")
     
     Case 4  ' salir
       Unload Me
  End Select
  
  Exit Sub
   
X:
  If Index = 2 And Err.Number = -2147217873 Then
    MsgBox "Registro no podrá Eliminarse mientras exista Información en la Tablas Relacionadas", vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & Err.Description & "  " & Err.Number, vbInformation, Caption
  End If
   
End Sub

Private Sub cAcepta_Click()
  Dim SQL As String
  Dim SQL1 As String
  'On Error GoTo X
  
  Set VGvardllgen = New dllgeneral.dll_general
  VGCNx.BeginTrans
  
  If modoinsert = True Then
    SQL = "INSERT INTO CT_ASIENTO (asientocodigo,asientodescripcion,flaggrabado,controlnref,nemotecref,"
    SQL = SQL & "zomames01,zomames02,zomames03,zomames04,zomames05,zomames06,zomames07,zomames08,zomames09,zomames10,zomames11,zomames12,usuariocodigo,fechaact,librocodigo,asientoadicionacargo) "
    SQL = SQL & "VALUES ('" & txt(0).Text & "','" & txt(1).Text & "'," & chk(0).Value & "," & chk(1).Value & ",'" & Trim$(UCase$(txt(2).Text)) & "',"
    SQL = SQL & VGvardllgen.ESNULO(txt(3).Text, 0) & "," & VGvardllgen.ESNULO(txt(4).Text, 0) & "," & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
    SQL = SQL & VGvardllgen.ESNULO(txt(6).Text, 0) & "," & VGvardllgen.ESNULO(txt(7).Text, 0) & "," & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
    SQL = SQL & VGvardllgen.ESNULO(txt(9).Text, 0) & "," & VGvardllgen.ESNULO(txt(10).Text, 0) & "," & VGvardllgen.ESNULO(txt(11).Text, 0) & "," & VGvardllgen.ESNULO(txt(12).Text, 0) & "," & VGvardllgen.ESNULO(txt(13).Text, 0) & "," & VGvardllgen.ESNULO(txt(14).Text, 0) & ",'"
    SQL = SQL & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "','" & IIf(Ctr_Ayuda1.xclave <> Empty, Trim$(Ctr_Ayuda1.xclave), "00") & "','" & ChkCargo.Value & "')"
    
    SQL1 = "INSERT INTO CT_ASIENTOCORRE (empresacodigo,asientocodigo,asientoanno,"
    SQL1 = SQL1 & "zomames01,zomames02,zomames03,zomames04,zomames05,zomames06,zomames07,zomames08,zomames09,zomames10,zomames11,zomames12,usuariocodigo,fechaact) "
    SQL1 = SQL1 & "VALUES ('" & VGParametros.empresacodigo & "',"
    SQL1 = SQL1 & "'" & txt(0).Text & " ','" & VGParamSistem.AnoProceso & "',"
    SQL1 = SQL1 & VGvardllgen.ESNULO(txt(3).Text, 0) & "," & VGvardllgen.ESNULO(txt(4).Text, 0) & "," & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
    SQL1 = SQL1 & VGvardllgen.ESNULO(txt(6).Text, 0) & "," & VGvardllgen.ESNULO(txt(7).Text, 0) & "," & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
    SQL1 = SQL1 & VGvardllgen.ESNULO(txt(9).Text, 0) & "," & VGvardllgen.ESNULO(txt(10).Text, 0) & "," & VGvardllgen.ESNULO(txt(11).Text, 0) & "," & VGvardllgen.ESNULO(txt(12).Text, 0) & "," & VGvardllgen.ESNULO(txt(13).Text, 0) & "," & VGvardllgen.ESNULO(txt(14).Text, 0) & ",'"
    SQL1 = SQL1 & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
    
                  
  ElseIf modoedit = True Then
    SQL = "UPDATE CT_ASIENTO SET asientodescripcion='" & Trim$(UCase$(txt(1).Text)) & "',"
    SQL = SQL & "flaggrabado=" & chk(0).Value & ","
    SQL = SQL & "controlnref=" & chk(1).Value & ","
    SQL = SQL & "nemotecref='" & txt(2).Text & "',"
    SQL = SQL & "zomames01=" & VGvardllgen.ESNULO(txt(3).Text, 0) & ",zomames02=" & VGvardllgen.ESNULO(txt(4).Text, 0) & ",zomames03=" & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
    SQL = SQL & "zomames04=" & VGvardllgen.ESNULO(txt(6).Text, 0) & ",zomames05=" & VGvardllgen.ESNULO(txt(7).Text, 0) & ",zomames06=" & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
    SQL = SQL & "zomames07=" & VGvardllgen.ESNULO(txt(9).Text, 0) & ",zomames08=" & VGvardllgen.ESNULO(txt(10).Text, 0) & ",zomames09=" & VGvardllgen.ESNULO(txt(11).Text, 0) & ","
    SQL = SQL & "zomames10=" & VGvardllgen.ESNULO(txt(12).Text, 0) & ",zomames11=" & VGvardllgen.ESNULO(txt(13).Text, 0) & ",zomames12=" & VGvardllgen.ESNULO(txt(14).Text, 0) & ","
    SQL = SQL & "usuariocodigo='" & VGusuario & "',fechaact='" & Format(Date, "dd/mm/yyyy") & "',"
    SQL = SQL & "librocodigo='" & IIf(Ctr_Ayuda1.xclave <> Empty, Trim$(Ctr_Ayuda1.xclave), "00") & "', "
    SQL = SQL & "asientoadicionacargo='" & ChkCargo.Value & "' "
    SQL = SQL & "WHERE asientocodigo='" & txt(0).Text & "'"
  
    SQL1 = "UPDATE CT_ASIENTOCORRE SET "
    SQL1 = SQL1 & "zomames01=" & VGvardllgen.ESNULO(txt(3).Text, 0) & "0,zomames02=" & VGvardllgen.ESNULO(txt(4).Text, 0) & ",zomames03=" & VGvardllgen.ESNULO(txt(5).Text, 0) & ","
    SQL1 = SQL1 & "zomames04=" & VGvardllgen.ESNULO(txt(6).Text, 0) & ",zomames05=" & VGvardllgen.ESNULO(txt(7).Text, 0) & ",zomames06=" & VGvardllgen.ESNULO(txt(8).Text, 0) & ","
    SQL1 = SQL1 & "zomames07=" & VGvardllgen.ESNULO(txt(9).Text, 0) & ",zomames08=" & VGvardllgen.ESNULO(txt(10).Text, 0) & ",zomames09=" & VGvardllgen.ESNULO(txt(11).Text, 0) & ","
    SQL1 = SQL1 & "zomames10=" & VGvardllgen.ESNULO(txt(12).Text, 0) & ",zomames11=" & VGvardllgen.ESNULO(txt(13).Text, 0) & ",zomames12=" & VGvardllgen.ESNULO(txt(14).Text, 0) & ","
    SQL1 = SQL1 & "usuariocodigo='" & VGusuario & "',fechaact='" & Format(Date, "dd/mm/yyyy") & "' "
    SQL1 = SQL1 & "WHERE empresacodigo='" & VGParametros.empresacodigo & "' and  asientocodigo='" & txt(0).Text & "' AND "
    SQL1 = SQL1 & "asientoanno='" & VGParamSistem.AnoProceso & "'"
  
  End If
  
  VGCNx.Execute (SQL)
  VGCNx.Execute (SQL1)
  VGCNx.CommitTrans
  
  Set VGvardllgen = Nothing
  frmbotones.Visible = True
  modoinsert = False: modoedit = False
  Call MuestraDatosAsiento
  cAcepta.Enabled = False
  Set VGvardllgen = Nothing
  Call ModoEditable(False, FrmVentasxMeses, "")
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  SSTab1.TabEnabled(2) = False
  Exit Sub

X:
  If Err.Number = -2147217873 Then
    MsgBox "Esta intentando registrar Código de Asiento Existente ", vbInformation, Caption
    txt(0).SetFocus
  Else
    MsgBox "Error inesperado: " & Err.Number & " " & Err.Description, vbInformation, Caption
  End If
  VGCNx.RollbackTrans
     
End Sub

Private Sub cCancela_Click()
  frmbotones.Visible = True
  modoinsert = False: modoedit = False
  cAcepta.Enabled = False
  SSTab1.Tab = 0
  SSTab1.TabEnabled(0) = True
  SSTab1.TabEnabled(1) = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If PreviousTab = 0 Then SSTab1.TabEnabled(PreviousTab) = False
End Sub


Private Sub TDBGridAsientos_DblClick()
    If rs.RecordCount > 0 Then Call cmdBotones_Click(1)
End Sub

'FIXIT: Declare 'LastRow' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Private Sub TDBGridAsientos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call EditarAsiento
End Sub

Private Sub txt_Change(Index As Integer)
 If modoinsert = True Or modoedit = True Then
   cAcepta.Enabled = ValidaDataIngreso()
 End If
End Sub

Private Sub chk_Click(Index As Integer)
    If modoinsert = True Or modoedit = True Then cAcepta.Enabled = ValidaDataIngreso()
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If Index = 2 And KeyCode = 13 Then
    cAcepta.SetFocus
 End If

End Sub

Private Sub txt_LostFocus(Index As Integer)
 If Index = 0 Then
   If Not IsNull(txt(0).Text) Then txt(0).Text = Format(txt(0).Text, "000")
 Else
   txt(Index).Text = UCase$(txt(Index).Text)
 End If
End Sub

Sub Editarfrilla()
 Dim i As Integer
 
 If rs.RecordCount > 0 Then
    With DataGrid1
        txt(0).Text = .Columns(0).Value
        For i = 3 To 14
             txt(i).Text = .Columns(i + 2).Value
        Next
    End With
 End If
End Sub

Sub ConfiguraGrid1()
 Dim i As Integer
 With DataGrid1
   .Columns(0).Width = 700
   .Columns(1).Width = 2500
 End With

End Sub

Function ValidaDataIngreso() As Boolean
 Dim i As Integer
  
  For i = 0 To 1
   If txt(i).Text = Empty Then
     ValidaDataIngreso = False
     Exit Function
   End If
   
  Next
  
  ValidaDataIngreso = True
End Function

