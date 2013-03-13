VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4D137D9C-00A6-4458-9B46-7E95DB76D55B}#9.0#0"; "TextFer.ocx"
Begin VB.Form frmMantSaldosInicialPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos Iniciales"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6270
   Begin TextFer.TxFer txtBuscar 
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   420
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   529
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
      Valor           =   ""
      NoCaracteres    =   "',*%"
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   360
      Index           =   1
      Left            =   3225
      TabIndex        =   7
      Top             =   4905
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Editar Saldo Iniciales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   15
      TabIndex        =   6
      Top             =   2895
      Width           =   6255
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   0
         Left            =   2880
         TabIndex        =   1
         Top             =   495
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   529
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
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Top             =   795
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   529
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
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   2
         Left            =   2895
         TabIndex        =   3
         Top             =   1185
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   529
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
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin TextFer.TxFer txt 
         Height          =   300
         Index           =   3
         Left            =   2895
         TabIndex        =   4
         Top             =   1485
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   529
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
         NumeroDecimales =   2
         SignoNegativo   =   0   'False
         MarcarTextoAlEnfoque=   -1  'True
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   75
         TabIndex        =   14
         Top             =   225
         Width           =   6135
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Haber (Dolares)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   75
         TabIndex        =   11
         Top             =   1500
         Width           =   2835
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Debe (Dolares)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   75
         TabIndex        =   10
         Top             =   1200
         Width           =   2835
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Haber (Soles)"
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   810
         Width           =   2805
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Debe (Soles)"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   510
         Width           =   2805
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   360
      Index           =   0
      Left            =   1665
      TabIndex        =   5
      Top             =   4905
      Width           =   1335
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   1860
      Left            =   15
      TabIndex        =   0
      Top             =   735
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   3281
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Código"
      Columns(0).DataField=   "cuentacodigo"
      Columns(0).DataWidth=   800
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción Plan Cuenta"
      Columns(1).DataField=   "cuentadescripcion"
      Columns(1).DataWidth=   2500
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=1,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H8000000D&,.bold=0"
      _StyleDefs(18)  =   ":id=6,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(19)  =   ":id=6,.fontname=MS Sans Serif"
      _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label3 
      Caption         =   "Buscar por Código / Descripción"
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   45
      TabIndex        =   13
      Top             =   165
      Width           =   3015
   End
End
Attribute VB_Name = "frmMantSaldosInicialPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim rsSaldo As ADODB.Recordset
Dim NombreTabla As String
Dim FlagGrabado As Boolean

Private Sub Form_Load()
 Dim SQL As String
 
 Set rs = New ADODB.Recordset
 Set rsSaldo = New ADODB.Recordset
 SQL = "SELECT cuentacodigo,cuentadescripcion FROM ct_cuenta WHERE empresacodigo='" & VGParametros.empresacodigo & "' and cuentanivel=" & VGnumnivelescuenta & " AND cuentacodigo<>'00' order by 1"
 Set rs = VGCNx.Execute(SQL)
 If rs.RecordCount <= 0 Then
    MsgBox "Faltan las Cuentas para los Saldos", vbInformation, Caption
    Exit Sub
 End If

 Set TDBGrid1.DataSource = rs
 Call Config_TDBGrid1
 NombreTabla = "CT_SALDOS" & VGParamSistem.Anoproceso
 Me.Height = 5865
 Me.Width = 6390
 cmdBotones(0).Enabled = False
 FlagGrabado = False
End Sub

Sub Editar()
    Dim SQL As String
    If rs.RecordCount > 0 Then
        SQL = "SELECT saldodebe00,saldohaber00,saldoussdebe00,saldousshaber00 "
        SQL = SQL & "FROM " & NombreTabla & " WHERE empresacodigo='" & VGParametros.empresacodigo & "' and cuentacodigo='" & TDBGrid1.Columns(0).Text & "'"
        Set rsSaldo = VGCNx.Execute(SQL)
        If rsSaldo.RecordCount > 0 Then
            txt(0).Text = rsSaldo(0)
            txt(1).Text = rsSaldo(1)
            txt(2).Text = rsSaldo(2)
            txt(3).Text = rsSaldo(3)
        Else
            txt(0).Text = 0: txt(1).Text = 0: txt(2).Text = 0: txt(3).Text = 0
        End If
      
        lbl(2).Caption = "Saldos de " & TDBGrid1.Columns(0).Text & " - " & TDBGrid1.Columns(1).Text
    End If
End Sub

Sub GrabarData()
    Dim rs As ADODB.Recordset
    Dim SQL As String
    Set rs = New ADODB.Recordset
    Set VGvardllgen = New dllgeneral.dll_general
    
    SQL = "SELECT cuentacodigo FROM " & NombreTabla & " WHERE empresacodigo='" & VGParametros.empresacodigo & "' and  cuentacodigo='" & TDBGrid1.Columns(0).Text & "'"
    If VGvardllgen.VerificaDatoExistente(VGCNx, SQL) > 0 Then
        SQL = "UPDATE " & NombreTabla & " SET "
        SQL = SQL & "saldodebe00=" & Val(VGvardllgen.ESNULO(txt(0).Text, 0)) & ","
        SQL = SQL & "saldohaber00=" & Val(VGvardllgen.ESNULO(txt(1).Text, 0)) & ","
        SQL = SQL & "saldoussdebe00=" & Val(VGvardllgen.ESNULO(txt(2).Text, 0)) & ","
        SQL = SQL & "saldousshaber00=" & Val(VGvardllgen.ESNULO(txt(3).Text, 0)) & ","
        SQL = SQL & "usuariocodigo='" & VGusuario & "',"
        SQL = SQL & "fechaact='" & Date & "' "
        SQL = SQL & "WHERE empresacodigo='" & VGParametros.empresacodigo & "' and cuentacodigo='" & TDBGrid1.Columns(0).Text & "'"
        VGCNx.Execute (SQL)
        FlagGrabado = True
    Else
        SQL = "INSERT " & NombreTabla & " (empresacodigo,cuentacodigo,saldodebe00,saldohaber00,saldoussdebe00,saldousshaber00,usuariocodigo,fechaact) "
        SQL = SQL & "VALUES ('" & VGParametros.empresacodigo & "','" & TDBGrid1.Columns(0).Text & "'," & Val(VGvardllgen.ESNULO(txt(0).Text, 0)) & "," & Val(VGvardllgen.ESNULO(txt(1).Text, 0)) & "," & Val(VGvardllgen.ESNULO(txt(2).Text, 0)) & ","
        SQL = SQL & Val(VGvardllgen.ESNULO(txt(3).Text, 0)) & ",'" & VGusuario & "','" & Date & "')"
        VGCNx.Execute (SQL)
        FlagGrabado = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FlagGrabado = True Then Call RecalcularAcumulados
End Sub

'FIXIT: Declare 'LastRow' con un tipo de datos de enlace en tiempo de compilación          FixIT90210ae-R1672-R1B8ZE
Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rs.RecordCount > 0 Then
       Call Editar
    End If
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Select Case Index
    Case 0:
       Call GrabarData
       cmdBotones(0).Enabled = False
    Case 1:
       Unload Me
       Set rs = Nothing
       Set rsSaldo = Nothing
  End Select
End Sub

Sub Config_TDBGrid1()
    TDBGrid1.Columns(0).Width = 900
    TDBGrid1.Columns(1).Width = 4800
End Sub

'FIXIT: Declare 'Buscar' con un tipo de datos de enlace en tiempo de compilación           FixIT90210ae-R1672-R1B8ZE
Function Buscar()
Dim SQL As String
  If IsNumeric(txtBuscar.Text) = True Then
    rs.Filter = "cuentacodigo like '" & Trim$(txtBuscar.Text) & "%'"
    Set TDBGrid1.DataSource = rs
  Else
    If Trim$(txtBuscar.Text) = Empty Then
        SQL = "select cuentacodigo,cuentadescripcion from ct_cuenta where cuentanivel=" & VGnumnivelescuenta & " AND cuentacodigo<>'00'"
        Set rs = VGCNx.Execute(SQL)
    Else
        rs.Filter = "cuentadescripcion like '%" & Trim$(txtBuscar.Text) & "%'"
    End If
    Set TDBGrid1.DataSource = rs
  End If
End Function

Private Sub txt_Change(Index As Integer)
  cmdBotones(0).Enabled = True
End Sub

Private Sub txtBuscar_Change()
  Call Buscar
End Sub

Sub RecalcularAcumulados()
  Set VGCommandoSP = New ADODB.Command
  VGCommandoSP.ActiveConnection = VGGeneral
  VGCommandoSP.CommandType = adCmdStoredProc
  VGCommandoSP.CommandText = "ct_recalacum_pro"
  VGCommandoSP.Parameters.Refresh
  With VGCommandoSP
        .Parameters("@base") = VGParamSistem.BDEmpresa
        .Parameters("@anno") = VGParamSistem.Anoproceso
        .Parameters("@empresa") = VGParametros.empresacodigo
        .Parameters("@mespro") = "01"
        .Parameters("@user") = VGParamSistem.Usuario
  End With
  VGCommandoSP.Execute

End Sub
